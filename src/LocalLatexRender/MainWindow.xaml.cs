using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.UI;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Text;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Media.Imaging;
using Microsoft.Web.WebView2.Core;
using Windows.ApplicationModel.DataTransfer;
using Windows.Graphics;
using Windows.Graphics.Imaging;
using Windows.Storage.Streams;
using WinRT.Interop;

namespace LocalLatexRender;

public sealed partial class MainWindow : Window
{
    private const string LocalHostName = "latex-render.local";
    private const int RenderViewportWidth = 2800;
    private const int RenderViewportHeight = 1600;
    private const int CropPadding = 24;
    private const byte AlphaThreshold = 8;
    private const int HistoryLimit = 8;
    private const int ClipboardBusyHResult = unchecked((int)0x800401D0);
    private const uint CfDibV5 = 17;
    private const uint BiBitFields = 3;
    private const uint LcsSrgb = 0x73524742;
    private const uint GmemMoveable = 0x0002;

    private static readonly Regex LatexEnvelopeRegex = new(
        @"^\s*(?<delimiter>\$\$|\$)(?<body>[\s\S]+?)\k<delimiter>\s*$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.Singleline);

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    private readonly SemaphoreSlim _clipboardGate = new(1, 1);
    private readonly SemaphoreSlim _manualRenderGate = new(1, 1);
    private readonly DispatcherQueueTimer _manualPreviewTimer;
    private Task? _startupTask;
    private bool _clipboardHooked;
    private bool _webViewReady;
    private InMemoryRandomAccessStream? _manualPreviewStream;
    private LatexPayload? _manualPreviewPayload;

    private ToggleSwitch ManualDisplayModeToggle = null!;
    private TextBox ManualInputTextBox = null!;
    private TextBlock ManualStatusTextBlock = null!;
    private TextBlock PreviewMetaTextBlock = null!;
    private Image ManualPreviewImage = null!;
    private StackPanel PreviewPlaceholderPanel = null!;
    private Button CopyManualButton = null!;

    public ObservableCollection<RenderHistoryItem> HistoryItems { get; } = [];

    public MainWindow()
    {
        InitializeComponent();
        BuildManualWorkbenchUi();

        _manualPreviewTimer = DispatcherQueue.CreateTimer();
        _manualPreviewTimer.Interval = TimeSpan.FromMilliseconds(280);
        _manualPreviewTimer.IsRepeating = false;
        _manualPreviewTimer.Tick += ManualPreviewTimer_Tick;
        RenderScaleComboBox.SelectionChanged += RenderScaleComboBox_SelectionChanged;

        Activated += MainWindow_Activated;
        Closed += MainWindow_Closed;

        SetManualPreviewEmptyState("输入公式后会自动刷新右侧实时反馈图片。");
    }

    private void BuildManualWorkbenchUi()
    {
        if (Content is not Grid rootGrid)
        {
            return;
        }

        var existingContentGrid = rootGrid.Children
            .OfType<Grid>()
            .FirstOrDefault(grid => Grid.GetRow(grid) == 2);

        if (existingContentGrid is null)
        {
            return;
        }

        var manualBorder = CreateCardBorder();
        manualBorder.Child = CreateManualWorkbenchContent();
        existingContentGrid.ColumnDefinitions.Clear();
        existingContentGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1.65, GridUnitType.Star) });
        existingContentGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(0.95, GridUnitType.Star) });
        existingContentGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var historyBorder = existingContentGrid.Children.OfType<Border>().FirstOrDefault();
        if (historyBorder is not null)
        {
            Grid.SetColumn(historyBorder, 1);
        }

        var existingCanvas = existingContentGrid.Children.OfType<Canvas>().FirstOrDefault();
        if (existingCanvas is not null)
        {
            Grid.SetColumn(existingCanvas, 2);
        }

        Grid.SetColumn(manualBorder, 0);
        existingContentGrid.Children.Insert(0, manualBorder);
    }

    private UIElement CreateManualWorkbenchContent()
    {
        var outerGrid = new Grid
        {
            RowSpacing = 16
        };
        outerGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        outerGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        outerGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

        var introPanel = new StackPanel
        {
            Spacing = 4
        };
        introPanel.Children.Add(CreateTitleTextBlock("手动公式工作台"));
        introPanel.Children.Add(CreateSecondaryTextBlock("支持直接输入原始 LaTeX，也支持粘贴 $...$ 或 $$...$$。"));
        outerGrid.Children.Add(introPanel);

        var actionsPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 12
        };
        ManualDisplayModeToggle = new ToggleSwitch
        {
            Header = "块级预览",
            IsOn = true,
            OffContent = "行内",
            OnContent = "块级"
        };
        ManualDisplayModeToggle.Toggled += ManualDisplayModeToggle_Toggled;

        var pasteButton = new Button
        {
            Content = "粘贴公式",
            MinWidth = 96,
            VerticalAlignment = VerticalAlignment.Bottom
        };
        pasteButton.Click += PasteManualButton_Click;

        CopyManualButton = new Button
        {
            Content = "复制结果",
            MinWidth = 96,
            IsEnabled = false,
            VerticalAlignment = VerticalAlignment.Bottom
        };
        CopyManualButton.Click += CopyManualButton_Click;

        actionsPanel.Children.Add(ManualDisplayModeToggle);
        actionsPanel.Children.Add(pasteButton);
        actionsPanel.Children.Add(CopyManualButton);
        Grid.SetRow(actionsPanel, 1);
        outerGrid.Children.Add(actionsPanel);

        var workbenchGrid = new Grid
        {
            ColumnSpacing = 16
        };
        workbenchGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1.1, GridUnitType.Star) });
        workbenchGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(0.9, GridUnitType.Star) });
        Grid.SetRow(workbenchGrid, 2);

        var inputGrid = new Grid
        {
            RowSpacing = 10
        };
        inputGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        inputGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

        ManualStatusTextBlock = CreateSecondaryTextBlock("输入公式后会自动刷新右侧实时反馈图片。");
        inputGrid.Children.Add(ManualStatusTextBlock);

        ManualInputTextBox = new TextBox
        {
            AcceptsReturn = true,
            FontFamily = new FontFamily("Cascadia Code"),
            PlaceholderText = "输入 LaTeX 公式",
            TextWrapping = TextWrapping.Wrap,
            MinHeight = 420
        };
        ManualInputTextBox.TextChanged += ManualInputTextBox_TextChanged;
        Grid.SetRow(ManualInputTextBox, 1);
        inputGrid.Children.Add(ManualInputTextBox);
        workbenchGrid.Children.Add(inputGrid);

        var previewShell = new Border
        {
            Padding = new Thickness(14),
            Background = GetThemeBrush("LayerFillColorDefaultBrush"),
            BorderBrush = GetThemeBrush("CardStrokeColorDefaultBrush"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(16)
        };
        Grid.SetColumn(previewShell, 1);

        var previewGrid = new Grid
        {
            RowSpacing = 10
        };
        previewGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        previewGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

        PreviewMetaTextBlock = CreateSecondaryTextBlock("等待渲染结果...");
        previewGrid.Children.Add(PreviewMetaTextBlock);

        var previewScroller = new ScrollViewer
        {
            HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        };
        Grid.SetRow(previewScroller, 1);

        var previewStage = new Grid
        {
            MinHeight = 320
        };

        ManualPreviewImage = new Image
        {
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            Stretch = Stretch.None,
            Visibility = Visibility.Collapsed
        };

        PreviewPlaceholderPanel = new StackPanel
        {
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            Spacing = 8
        };
        PreviewPlaceholderPanel.Children.Add(CreateTitleTextBlock("实时反馈图片区", 18));
        PreviewPlaceholderPanel.Children.Add(new TextBlock
        {
            Text = "输入有效公式后，这里会显示透明 PNG 预览。",
            TextAlignment = TextAlignment.Center,
            TextWrapping = TextWrapping.Wrap,
            HorizontalAlignment = HorizontalAlignment.Center,
            Foreground = GetThemeBrush("TextFillColorSecondaryBrush")
        });

        previewStage.Children.Add(ManualPreviewImage);
        previewStage.Children.Add(PreviewPlaceholderPanel);
        previewScroller.Content = previewStage;
        previewGrid.Children.Add(previewScroller);
        previewShell.Child = previewGrid;
        workbenchGrid.Children.Add(previewShell);

        outerGrid.Children.Add(workbenchGrid);
        return outerGrid;
    }

    private Border CreateCardBorder()
    {
        return new Border
        {
            Padding = new Thickness(20),
            Background = GetThemeBrush("CardBackgroundFillColorDefaultBrush"),
            BorderBrush = GetThemeBrush("CardStrokeColorDefaultBrush"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(20)
        };
    }

    private TextBlock CreateTitleTextBlock(string text, double fontSize = 18)
    {
        return new TextBlock
        {
            Text = text,
            FontSize = fontSize,
            FontWeight = FontWeights.SemiBold
        };
    }

    private TextBlock CreateSecondaryTextBlock(string text)
    {
        return new TextBlock
        {
            Text = text,
            TextWrapping = TextWrapping.WrapWholeWords,
            Foreground = GetThemeBrush("TextFillColorSecondaryBrush")
        };
    }

    private Brush? GetThemeBrush(string resourceKey)
    {
        return Application.Current.Resources.TryGetValue(resourceKey, out var resource)
            ? resource as Brush
            : null;
    }

    private void MainWindow_Activated(object sender, WindowActivatedEventArgs args)
    {
        _startupTask ??= InitializeAsync();
    }

    private async Task InitializeAsync()
    {
        try
        {
            DebugLog("开始初始化窗口。");
            ConfigureWindow();
            UpdateStatus("正在加载本地 KaTeX 渲染器...");
            UpdateManualStatus("正在准备手动预览模块...");

            await InitializeWebViewAsync();

            Clipboard.ContentChanged += Clipboard_ContentChanged;
            _clipboardHooked = true;

            DebugLog("初始化完成，已挂接 Clipboard.ContentChanged。");
            UpdateStatus("自动转换已开启，等待剪贴板中的 LaTeX 公式。");
            UpdateManualStatus("手动工作台已就绪，支持实时预览。");
            ScheduleManualPreview();
        }
        catch (Exception ex)
        {
            DebugLog($"初始化失败：{ex}");
            UpdateStatus($"初始化失败：{ex.Message}");
            UpdateManualStatus($"初始化失败：{ex.Message}");
        }
    }

    private void ConfigureWindow()
    {
        Title = "轻量 LaTeX 剪贴板渲染器";
        RendererWebView.Width = RenderViewportWidth;
        RendererWebView.Height = RenderViewportHeight;

        var hwnd = WindowNative.GetWindowHandle(this);
        var windowId = Win32Interop.GetWindowIdFromWindow(hwnd);
        var appWindow = AppWindow.GetFromWindowId(windowId);
        appWindow.Resize(new SizeInt32(1360, 860));
    }

    private async Task InitializeWebViewAsync()
    {
        var webRoot = Path.Combine(AppContext.BaseDirectory, "Web");
        if (!Directory.Exists(webRoot))
        {
            throw new DirectoryNotFoundException($"未找到 Web 资源目录：{webRoot}");
        }

        await RendererWebView.EnsureCoreWebView2Async();
        RendererWebView.DefaultBackgroundColor = ColorHelper.FromArgb(0, 0, 0, 0);

        var settings = RendererWebView.CoreWebView2.Settings;
        settings.AreDevToolsEnabled = false;
        settings.AreDefaultContextMenusEnabled = false;
        settings.AreDefaultScriptDialogsEnabled = false;
        settings.IsStatusBarEnabled = false;
        settings.IsZoomControlEnabled = false;

        RendererWebView.CoreWebView2.SetVirtualHostNameToFolderMapping(
            LocalHostName,
            webRoot,
            CoreWebView2HostResourceAccessKind.Allow);

        var navigationTask = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        void HandleNavigationCompleted(object? _, CoreWebView2NavigationCompletedEventArgs eventArgs)
        {
            RendererWebView.CoreWebView2.NavigationCompleted -= HandleNavigationCompleted;
            if (eventArgs.IsSuccess)
            {
                navigationTask.TrySetResult(true);
                return;
            }

            navigationTask.TrySetException(
                new InvalidOperationException($"渲染页加载失败，WebErrorStatus={eventArgs.WebErrorStatus}"));
        }

        RendererWebView.CoreWebView2.NavigationCompleted += HandleNavigationCompleted;
        RendererWebView.Source = new Uri($"https://{LocalHostName}/index.html");
        await navigationTask.Task;

        _webViewReady = await WaitForRenderBridgeAsync();
        if (!_webViewReady)
        {
            throw new InvalidOperationException("渲染页桥接脚本没有正确初始化。");
        }

        DebugLog("WebView2 渲染桥接已就绪。");
    }

    private async Task<bool> WaitForRenderBridgeAsync()
    {
        const string readinessScript = "Boolean(window.renderBridge && typeof window.renderBridge.render === 'function')";

        for (var attempt = 0; attempt < 20; attempt++)
        {
            var rawResult = await RendererWebView.CoreWebView2.ExecuteScriptAsync(readinessScript);
            DebugLog($"桥接检测第 {attempt + 1} 次：{rawResult}");
            if (rawResult.Contains("true", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            await Task.Delay(100);
        }

        return false;
    }

    private void Clipboard_ContentChanged(object? sender, object e)
    {
        DebugLog("收到 Clipboard.ContentChanged 事件。");
        if (!DispatcherQueue.TryEnqueue(() => _ = ProcessClipboardAsync()))
        {
            _ = ProcessClipboardAsync();
        }
    }

    private async Task ProcessClipboardAsync()
    {
        if (!_webViewReady || !AutoConvertToggle.IsOn)
        {
            return;
        }

        if (!await _clipboardGate.WaitAsync(0))
        {
            return;
        }

        try
        {
            var clipboardContent = Clipboard.GetContent();
            if (!clipboardContent.Contains(StandardDataFormats.Text))
            {
                DebugLog("剪贴板当前不包含文本格式。");
                return;
            }

            var clipboardText = await clipboardContent.GetTextAsync();
            if (string.IsNullOrWhiteSpace(clipboardText))
            {
                DebugLog("剪贴板文本为空。");
                return;
            }

            DebugLog($"读取到剪贴板文本：{clipboardText}");

            if (!TryParseLatexEnvelope(clipboardText, out var latexPayload))
            {
                DebugLog("文本未命中 LaTeX 包裹规则。");
                return;
            }

            var renderScale = GetSelectedRenderScale();
            UpdateStatus($"已检测到公式，正在按 {renderScale}x 输出渲染透明 PNG...");
            DebugLog($"开始渲染公式：{latexPayload.Expression}");

            var renderResult = await RenderLatexToImageAsync(latexPayload.Expression, latexPayload.DisplayMode, renderScale);
            if (!renderResult.Ok || renderResult.ImageStream is null)
            {
                DebugLog($"渲染失败：{renderResult.ErrorMessage}");
                UpdateStatus($"渲染失败：{renderResult.ErrorMessage ?? "KaTeX 返回了未知错误。"}");
                return;
            }

            await WriteClipboardImageAsync(renderResult.ImageStream);
            DebugLog("图片已回写到剪贴板。");

            AddHistory(latexPayload.OriginalText, latexPayload.DisplayMode, "自动");
            UpdateStatus(latexPayload.DisplayMode
                ? $"已转换块级公式，并以 {renderScale}x 输出写回剪贴板。"
                : $"已转换行内公式，并以 {renderScale}x 输出写回剪贴板。");
        }
        catch (Exception ex)
        {
            DebugLog($"处理剪贴板时出错：{ex}");
            UpdateStatus(ex is COMException { HResult: ClipboardBusyHResult }
                ? "系统剪贴板正忙，后台转换会自动在下一次复制时继续尝试。"
                : $"处理剪贴板时出错：{ex.Message}");
        }
        finally
        {
            _clipboardGate.Release();
        }
    }

    private static bool TryParseLatexEnvelope(string clipboardText, out LatexPayload payload)
    {
        payload = default;
        var match = LatexEnvelopeRegex.Match(clipboardText);
        if (!match.Success)
        {
            return false;
        }

        var body = match.Groups["body"].Value.Trim();
        if (string.IsNullOrWhiteSpace(body))
        {
            return false;
        }

        payload = new LatexPayload(
            clipboardText,
            body,
            string.Equals(match.Groups["delimiter"].Value, "$$", StringComparison.Ordinal));

        return true;
    }

    private static bool TryBuildManualPayload(string inputText, bool preferDisplayMode, out LatexPayload payload)
    {
        payload = default;

        if (string.IsNullOrWhiteSpace(inputText))
        {
            return false;
        }

        var trimmed = inputText.Trim();
        if (TryParseLatexEnvelope(trimmed, out payload))
        {
            return true;
        }

        payload = new LatexPayload(
            preferDisplayMode ? $"$${trimmed}$$" : $"${trimmed}$",
            trimmed,
            preferDisplayMode);

        return true;
    }

    private void ManualInputTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        ScheduleManualPreview();
    }

    private void ManualDisplayModeToggle_Toggled(object sender, RoutedEventArgs e)
    {
        ScheduleManualPreview();
    }

    private void RenderScaleComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_webViewReady)
        {
            UpdateStatus($"输出缩放已切换到 {GetSelectedRenderScale()}x。");
        }

        ScheduleManualPreview();
    }

    private int GetSelectedRenderScale()
    {
        if (RenderScaleComboBox.SelectedItem is ComboBoxItem comboBoxItem &&
            int.TryParse(comboBoxItem.Tag?.ToString(), out var scale))
        {
            return Math.Clamp(scale, 2, 5);
        }

        return 2;
    }

    private void ScheduleManualPreview()
    {
        _manualPreviewTimer.Stop();
        _manualPreviewTimer.Start();
    }

    private async void ManualPreviewTimer_Tick(DispatcherQueueTimer sender, object args)
    {
        sender.Stop();
        await RenderManualPreviewAsync();
    }

    private async Task RenderManualPreviewAsync()
    {
        if (!_webViewReady)
        {
            UpdateManualStatus("渲染器尚未就绪，稍后会自动刷新。");
            return;
        }

        await _manualRenderGate.WaitAsync();
        try
        {
            if (!TryBuildManualPayload(ManualInputTextBox.Text, ManualDisplayModeToggle.IsOn, out var payload))
            {
                _manualPreviewPayload = null;
                _manualPreviewStream?.Dispose();
                _manualPreviewStream = null;
                SetManualPreviewEmptyState("请输入公式，支持直接输入原始 LaTeX 或粘贴 $...$ / $$...$$。");
                return;
            }

            UpdateManualStatus("正在更新实时反馈图片...");
            PreviewMetaTextBlock.Text = "正在渲染透明 PNG 预览...";
            DebugLog($"开始更新手动预览：{payload.Expression}");

            var renderScale = GetSelectedRenderScale();
            var renderResult = await RenderLatexToImageAsync(payload.Expression, payload.DisplayMode, renderScale);
            if (!renderResult.Ok || renderResult.ImageStream is null)
            {
                _manualPreviewPayload = null;
                _manualPreviewStream?.Dispose();
                _manualPreviewStream = null;
                SetManualPreviewEmptyState($"预览渲染失败：{renderResult.ErrorMessage ?? "KaTeX 返回了未知错误。"}");
                return;
            }

            var imageSize = await GetImageSizeAsync(renderResult.ImageStream.CloneStream());
            var previewImageStream = await CloneStreamAsync(renderResult.ImageStream);
            var clipboardImageStream = await CloneStreamAsync(renderResult.ImageStream);

            await ShowManualPreviewAsync(previewImageStream);

            _manualPreviewStream?.Dispose();
            _manualPreviewStream = clipboardImageStream;
            _manualPreviewPayload = payload;

            PreviewMetaTextBlock.Text =
                $"{(payload.DisplayMode ? "块级" : "行内")} PNG 约 {imageSize.Width} × {imageSize.Height} 像素 · {renderScale}x 输出";
            DebugLog($"手动预览已更新：{payload.OriginalText}");
            UpdateManualStatus(payload.DisplayMode
                ? $"块级公式实时反馈图片已更新，当前为 {renderScale}x 输出。"
                : $"行内公式实时反馈图片已更新，当前为 {renderScale}x 输出。");
            CopyManualButton.IsEnabled = true;
        }
        catch (Exception ex)
        {
            DebugLog($"更新手动预览失败：{ex}");
            SetManualPreviewEmptyState($"预览更新失败：{ex.Message}");
        }
        finally
        {
            _manualRenderGate.Release();
        }
    }

    private async void PasteManualButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var clipboardContent = Clipboard.GetContent();
            if (!clipboardContent.Contains(StandardDataFormats.Text))
            {
                UpdateManualStatus("剪贴板里没有可粘贴的文本。");
                return;
            }

            var clipboardText = await clipboardContent.GetTextAsync();
            if (string.IsNullOrWhiteSpace(clipboardText))
            {
                UpdateManualStatus("剪贴板里的文本为空。");
                return;
            }

            ManualInputTextBox.Text = clipboardText;
            ManualInputTextBox.Focus(FocusState.Programmatic);
            DebugLog($"手动输入区已粘贴文本：{clipboardText}");
            UpdateManualStatus("已从剪贴板粘贴公式文本。");
            ScheduleManualPreview();
        }
        catch (Exception ex)
        {
            DebugLog($"粘贴到手动输入区失败：{ex}");
            UpdateManualStatus($"粘贴失败：{ex.Message}");
        }
    }

    private async void CopyManualButton_Click(object sender, RoutedEventArgs e)
    {
        if (_manualPreviewPayload is null || _manualPreviewStream is null)
        {
            UpdateManualStatus("当前还没有可复制的渲染结果。");
            return;
        }

        try
        {
            var clipboardStream = await CloneStreamAsync(_manualPreviewStream);
            await WriteClipboardImageAsync(clipboardStream);
            AddHistory(_manualPreviewPayload.Value.OriginalText, _manualPreviewPayload.Value.DisplayMode, "手动");
            DebugLog($"手动复制已完成：{_manualPreviewPayload.Value.OriginalText}");
            UpdateManualStatus("已把当前预览图片复制到剪贴板。");
        }
        catch (COMException ex) when (ex.HResult == ClipboardBusyHResult)
        {
            DebugLog($"复制手动结果失败：{ex}");
            UpdateManualStatus("系统剪贴板当前正忙，请稍等一秒后再点一次“复制结果”。");
        }
        catch (Exception ex)
        {
            DebugLog($"复制手动结果失败：{ex}");
            UpdateManualStatus($"复制失败：{ex.Message}");
        }
    }

    private async Task<RenderedImageResult> RenderLatexToImageAsync(string expression, bool displayMode, int scale)
    {
        var script = $"window.renderBridge.render({JsonSerializer.Serialize(new RenderRequest(expression, displayMode, scale), JsonOptions)})";
        var rawResult = await RendererWebView.CoreWebView2.ExecuteScriptAsync(script);
        DebugLog($"渲染脚本返回：{rawResult}");

        var renderResult = JsonSerializer.Deserialize<RenderResponse>(rawResult, JsonOptions);
        if (renderResult is null)
        {
            return new RenderedImageResult(false, null, "无法解析来自 WebView2 的返回结果。");
        }

        if (!renderResult.Ok)
        {
            return new RenderedImageResult(false, null, renderResult.Error ?? "KaTeX 渲染失败。");
        }

        await Task.Delay(120);

        var previewStream = new InMemoryRandomAccessStream();
        await RendererWebView.CoreWebView2.CapturePreviewAsync(
            CoreWebView2CapturePreviewImageFormat.Png,
            previewStream);

        var croppedStream = await CropTransparentEdgesAsync(previewStream, CropPadding);
        return new RenderedImageResult(true, croppedStream, null);
    }

    private static async Task<InMemoryRandomAccessStream> CropTransparentEdgesAsync(IRandomAccessStream sourceStream, int padding)
    {
        sourceStream.Seek(0);

        var decoder = await BitmapDecoder.CreateAsync(sourceStream);
        var pixelProvider = await decoder.GetPixelDataAsync(
            BitmapPixelFormat.Bgra8,
            BitmapAlphaMode.Ignore,
            new BitmapTransform(),
            ExifOrientationMode.IgnoreExifOrientation,
            ColorManagementMode.DoNotColorManage);

        var capturedPixels = pixelProvider.DetachPixelData();
        var width = (int)decoder.PixelWidth;
        var height = (int)decoder.PixelHeight;
        var pixels = new byte[capturedPixels.Length];

        var minX = width;
        var minY = height;
        var maxX = -1;
        var maxY = -1;

        for (var y = 0; y < height; y++)
        {
            for (var x = 0; x < width; x++)
            {
                var pixelOffset = ((y * width) + x) * 4;
                var alpha = DeriveAlphaFromOpaqueCapture(
                    capturedPixels[pixelOffset],
                    capturedPixels[pixelOffset + 1],
                    capturedPixels[pixelOffset + 2]);

                pixels[pixelOffset] = 0;
                pixels[pixelOffset + 1] = 0;
                pixels[pixelOffset + 2] = 0;
                pixels[pixelOffset + 3] = alpha;

                if (alpha <= AlphaThreshold)
                {
                    continue;
                }

                minX = Math.Min(minX, x);
                minY = Math.Min(minY, y);
                maxX = Math.Max(maxX, x);
                maxY = Math.Max(maxY, y);
            }
        }

        if (maxX < 0 || maxY < 0)
        {
            return await EncodePngAsync(new byte[4], 1, 1, decoder.DpiX, decoder.DpiY);
        }

        minX = Math.Max(0, minX - padding);
        minY = Math.Max(0, minY - padding);
        maxX = Math.Min(width - 1, maxX + padding);
        maxY = Math.Min(height - 1, maxY + padding);

        var croppedWidth = maxX - minX + 1;
        var croppedHeight = maxY - minY + 1;
        var croppedPixels = new byte[croppedWidth * croppedHeight * 4];

        for (var y = 0; y < croppedHeight; y++)
        {
            var sourceOffset = ((minY + y) * width + minX) * 4;
            var destinationOffset = y * croppedWidth * 4;
            System.Buffer.BlockCopy(pixels, sourceOffset, croppedPixels, destinationOffset, croppedWidth * 4);
        }

        return await EncodePngAsync(croppedPixels, croppedWidth, croppedHeight, decoder.DpiX, decoder.DpiY);
    }

    // WebView2 preview capture flattens transparency, so we render white-on-black
    // and reconstruct a transparent alpha mask from the captured brightness.
    private static byte DeriveAlphaFromOpaqueCapture(byte blue, byte green, byte red)
    {
        return Math.Max(red, Math.Max(green, blue));
    }

    private static async Task<InMemoryRandomAccessStream> EncodePngAsync(
        byte[] pixels,
        int width,
        int height,
        double dpiX,
        double dpiY)
    {
        var outputStream = new InMemoryRandomAccessStream();
        var encoder = await BitmapEncoder.CreateAsync(BitmapEncoder.PngEncoderId, outputStream);
        encoder.SetPixelData(
            BitmapPixelFormat.Bgra8,
            BitmapAlphaMode.Straight,
            (uint)width,
            (uint)height,
            dpiX,
            dpiY,
            pixels);

        await encoder.FlushAsync();
        outputStream.Seek(0);
        return outputStream;
    }

    private async Task WriteClipboardImageAsync(InMemoryRandomAccessStream imageStream)
    {
        imageStream.Seek(0);

        try
        {
            await WriteClipboardImageNativeAsync(imageStream.CloneStream());
            return;
        }
        catch (Exception ex)
        {
            DebugLog($"使用原生剪贴板格式写入失败，将回退到 DataPackage：{ex}");
        }

        var dataPackage = new DataPackage
        {
            RequestedOperation = DataPackageOperation.Copy
        };

        dataPackage.SetBitmap(RandomAccessStreamReference.CreateFromStream(imageStream.CloneStream()));

        for (var attempt = 0; attempt < 20; attempt++)
        {
            try
            {
                Clipboard.SetContent(dataPackage);
                Clipboard.Flush();
                return;
            }
            catch (COMException ex) when (ex.HResult == ClipboardBusyHResult && attempt < 19)
            {
                DebugLog($"剪贴板暂时被占用，第 {attempt + 1} 次重试。");
                await Task.Delay(150);
            }
        }
    }

    private async Task WriteClipboardImageNativeAsync(IRandomAccessStream imageStream)
    {
        imageStream.Seek(0);

        var pngBytes = await ReadAllBytesAsync(imageStream.CloneStream());
        var (dibBytes, dibFormatName) = await CreateClipboardDibAsync(imageStream.CloneStream());
        var pngFormat = RegisterClipboardFormat("PNG");
        if (pngFormat == 0)
        {
            throw new COMException("无法注册 PNG 剪贴板格式。", Marshal.GetHRForLastWin32Error());
        }

        for (var attempt = 0; attempt < 20; attempt++)
        {
            IntPtr pngHandle = IntPtr.Zero;
            IntPtr dibHandle = IntPtr.Zero;
            var clipboardOpened = false;

            try
            {
                pngHandle = CreateMoveableHGlobal(pngBytes);
                dibHandle = CreateMoveableHGlobal(dibBytes);

                clipboardOpened = OpenClipboard(WindowNative.GetWindowHandle(this));
                if (!clipboardOpened)
                {
                    throw new COMException("无法打开系统剪贴板。", Marshal.GetHRForLastWin32Error());
                }

                if (!EmptyClipboard())
                {
                    throw new COMException("无法清空系统剪贴板。", Marshal.GetHRForLastWin32Error());
                }

                if (SetClipboardData(pngFormat, pngHandle) == IntPtr.Zero)
                {
                    throw new COMException("写入 PNG 剪贴板格式失败。", Marshal.GetHRForLastWin32Error());
                }
                pngHandle = IntPtr.Zero;

                if (SetClipboardData(CfDibV5, dibHandle) == IntPtr.Zero)
                {
                    throw new COMException($"写入 {dibFormatName} 剪贴板格式失败。", Marshal.GetHRForLastWin32Error());
                }
                dibHandle = IntPtr.Zero;

                return;
            }
            catch (COMException ex) when (attempt < 19)
            {
                DebugLog($"原生剪贴板写入失败，第 {attempt + 1} 次重试：0x{ex.HResult:X8}");
                await Task.Delay(150);
            }
            finally
            {
                if (clipboardOpened)
                {
                    CloseClipboard();
                }

                if (pngHandle != IntPtr.Zero)
                {
                    GlobalFree(pngHandle);
                }

                if (dibHandle != IntPtr.Zero)
                {
                    GlobalFree(dibHandle);
                }
            }
        }

        throw new COMException("多次尝试后仍无法写入原生剪贴板格式。", ClipboardBusyHResult);
    }

    private static async Task<(byte[] DibBytes, string FormatName)> CreateClipboardDibAsync(IRandomAccessStream imageStream)
    {
        imageStream.Seek(0);

        var decoder = await BitmapDecoder.CreateAsync(imageStream);
        var pixelProvider = await decoder.GetPixelDataAsync(
            BitmapPixelFormat.Bgra8,
            BitmapAlphaMode.Straight,
            new BitmapTransform(),
            ExifOrientationMode.IgnoreExifOrientation,
            ColorManagementMode.DoNotColorManage);

        var pixels = pixelProvider.DetachPixelData();
        var dibPixels = new byte[pixels.Length];
        var width = (int)decoder.PixelWidth;
        var height = (int)decoder.PixelHeight;
        var xPelsPerMeter = (int)Math.Round(decoder.DpiX / 0.0254);
        var yPelsPerMeter = (int)Math.Round(decoder.DpiY / 0.0254);

        for (var pixelOffset = 0; pixelOffset < pixels.Length; pixelOffset += 4)
        {
            var alpha = pixels[pixelOffset + 3];
            dibPixels[pixelOffset] = CompositeChannelOnWhite(pixels[pixelOffset], alpha);
            dibPixels[pixelOffset + 1] = CompositeChannelOnWhite(pixels[pixelOffset + 1], alpha);
            dibPixels[pixelOffset + 2] = CompositeChannelOnWhite(pixels[pixelOffset + 2], alpha);
            dibPixels[pixelOffset + 3] = 255;
        }

        var header = new BitmapV5Header
        {
            Size = (uint)Marshal.SizeOf<BitmapV5Header>(),
            Width = width,
            Height = -height,
            Planes = 1,
            BitCount = 32,
            Compression = BiBitFields,
            SizeImage = (uint)dibPixels.Length,
            XPelsPerMeter = xPelsPerMeter,
            YPelsPerMeter = yPelsPerMeter,
            RedMask = 0x00FF0000,
            GreenMask = 0x0000FF00,
            BlueMask = 0x000000FF,
            AlphaMask = 0xFF000000,
            ColorSpaceType = LcsSrgb
        };

        var headerBytes = StructureToBytes(header);
        var dibBytes = new byte[headerBytes.Length + dibPixels.Length];
        System.Buffer.BlockCopy(headerBytes, 0, dibBytes, 0, headerBytes.Length);
        System.Buffer.BlockCopy(dibPixels, 0, dibBytes, headerBytes.Length, dibPixels.Length);

        return (dibBytes, "CF_DIBV5");
    }

    private static byte CompositeChannelOnWhite(byte channel, byte alpha)
    {
        return (byte)((channel * alpha + 255 * (255 - alpha) + 127) / 255);
    }

    private static byte[] StructureToBytes<T>(T value) where T : struct
    {
        var size = Marshal.SizeOf<T>();
        var bytes = new byte[size];
        var buffer = Marshal.AllocHGlobal(size);

        try
        {
            Marshal.StructureToPtr(value, buffer, false);
            Marshal.Copy(buffer, bytes, 0, size);
            return bytes;
        }
        finally
        {
            Marshal.FreeHGlobal(buffer);
        }
    }

    private static async Task<byte[]> ReadAllBytesAsync(IRandomAccessStream stream)
    {
        stream.Seek(0);

        using var inputStream = stream.GetInputStreamAt(0);
        using var reader = new DataReader(inputStream);
        var size = checked((uint)stream.Size);
        await reader.LoadAsync(size);

        var bytes = new byte[size];
        reader.ReadBytes(bytes);
        return bytes;
    }

    private static IntPtr CreateMoveableHGlobal(byte[] bytes)
    {
        var handle = GlobalAlloc(GmemMoveable, (nuint)bytes.Length);
        if (handle == IntPtr.Zero)
        {
            throw new OutOfMemoryException("无法为剪贴板数据分配内存。");
        }

        var locked = GlobalLock(handle);
        if (locked == IntPtr.Zero)
        {
            GlobalFree(handle);
            throw new COMException("无法锁定剪贴板内存。", Marshal.GetHRForLastWin32Error());
        }

        try
        {
            Marshal.Copy(bytes, 0, locked, bytes.Length);
            return handle;
        }
        finally
        {
            GlobalUnlock(handle);
        }
    }

    private async Task ShowManualPreviewAsync(IRandomAccessStream imageStream)
    {
        imageStream.Seek(0);
        var bitmapImage = new BitmapImage();
        await bitmapImage.SetSourceAsync(imageStream);

        ManualPreviewImage.Source = bitmapImage;
        ManualPreviewImage.Visibility = Visibility.Visible;
        PreviewPlaceholderPanel.Visibility = Visibility.Collapsed;
    }

    private void SetManualPreviewEmptyState(string message)
    {
        ManualPreviewImage.Source = null;
        ManualPreviewImage.Visibility = Visibility.Collapsed;
        PreviewPlaceholderPanel.Visibility = Visibility.Visible;
        PreviewMetaTextBlock.Text = message;
        UpdateManualStatus(message);
        CopyManualButton.IsEnabled = false;
    }

    private void UpdateStatus(string message)
    {
        StatusTextBlock.Text = $"{DateTimeOffset.Now:HH:mm:ss}  {message}";
    }

    private void UpdateManualStatus(string message)
    {
        ManualStatusTextBlock.Text = $"{DateTimeOffset.Now:HH:mm:ss}  {message}";
    }

    private void AddHistory(string originalText, bool displayMode, string sourceLabel)
    {
        HistoryItems.Insert(0, new RenderHistoryItem(
            originalText,
            $"{sourceLabel} · {(displayMode ? "块级公式" : "行内公式")}",
            DateTimeOffset.Now.ToString("HH:mm:ss")));

        while (HistoryItems.Count > HistoryLimit)
        {
            HistoryItems.RemoveAt(HistoryItems.Count - 1);
        }
    }

    private static async Task<InMemoryRandomAccessStream> CloneStreamAsync(IRandomAccessStream source)
    {
        var input = source.CloneStream();
        input.Seek(0);

        var clone = new InMemoryRandomAccessStream();
        await RandomAccessStream.CopyAsync(input, clone);
        clone.Seek(0);
        return clone;
    }

    private static async Task<ImageDimensions> GetImageSizeAsync(IRandomAccessStream source)
    {
        source.Seek(0);
        var decoder = await BitmapDecoder.CreateAsync(source);
        return new ImageDimensions(decoder.PixelWidth, decoder.PixelHeight);
    }

    private void AutoConvertToggle_Toggled(object sender, RoutedEventArgs e)
    {
        UpdateStatus(AutoConvertToggle.IsOn
            ? (_webViewReady ? "自动转换已开启，等待剪贴板中的 LaTeX 公式。" : "自动转换已开启，正在初始化渲染器...")
            : "自动转换已关闭。");
    }

    private void MainWindow_Closed(object sender, WindowEventArgs args)
    {
        if (_clipboardHooked)
        {
            Clipboard.ContentChanged -= Clipboard_ContentChanged;
        }

        _manualPreviewTimer.Tick -= ManualPreviewTimer_Tick;
        _manualPreviewStream?.Dispose();
        _clipboardGate.Dispose();
        _manualRenderGate.Dispose();
    }

    [Conditional("DEBUG")]
    private static void DebugLog(string message)
    {
        try
        {
            File.AppendAllText(
                Path.Combine(AppContext.BaseDirectory, "clipboard-render.log"),
                $"{DateTimeOffset.Now:HH:mm:ss.fff} {message}{Environment.NewLine}");
        }
        catch
        {
            // 调试日志不能影响主流程。
        }
    }

    public sealed class RenderHistoryItem
    {
        public RenderHistoryItem(string originalText, string modeText, string timestampText)
        {
            OriginalText = originalText;
            ModeText = modeText;
            TimestampText = timestampText;
        }

        public string OriginalText { get; }

        public string ModeText { get; }

        public string TimestampText { get; }
    }

    private readonly record struct LatexPayload(string OriginalText, string Expression, bool DisplayMode);

    private sealed record RenderRequest(string Expression, bool DisplayMode, int Scale);

    private sealed class RenderResponse
    {
        public bool Ok { get; init; }

        public string? Error { get; init; }
    }

    private readonly record struct ImageDimensions(uint Width, uint Height);

    private sealed record RenderedImageResult(bool Ok, InMemoryRandomAccessStream? ImageStream, string? ErrorMessage);

    [StructLayout(LayoutKind.Sequential)]
    private struct CieXyz
    {
        public int X;
        public int Y;
        public int Z;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct CieXyzTriple
    {
        public CieXyz Red;
        public CieXyz Green;
        public CieXyz Blue;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct BitmapV5Header
    {
        public uint Size;
        public int Width;
        public int Height;
        public ushort Planes;
        public ushort BitCount;
        public uint Compression;
        public uint SizeImage;
        public int XPelsPerMeter;
        public int YPelsPerMeter;
        public uint ClrUsed;
        public uint ClrImportant;
        public uint RedMask;
        public uint GreenMask;
        public uint BlueMask;
        public uint AlphaMask;
        public uint ColorSpaceType;
        public CieXyzTriple Endpoints;
        public uint GammaRed;
        public uint GammaGreen;
        public uint GammaBlue;
        public uint Intent;
        public uint ProfileData;
        public uint ProfileSize;
        public uint Reserved;
    }

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern uint RegisterClipboardFormat(string format);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool OpenClipboard(IntPtr newOwner);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CloseClipboard();

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool EmptyClipboard();

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetClipboardData(uint format, IntPtr memoryHandle);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr GlobalAlloc(uint flags, nuint bytes);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr GlobalLock(IntPtr handle);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GlobalUnlock(IntPtr handle);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr GlobalFree(IntPtr handle);
}
