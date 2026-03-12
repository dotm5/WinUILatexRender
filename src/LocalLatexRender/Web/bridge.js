const equationElement = document.getElementById("equation");

window.renderBridge = {
    ping() {
        return true;
    },

    render(payload) {
        const scale = Math.min(5, Math.max(2, Number(payload.Scale) || 2));

        equationElement.innerHTML = "";
        equationElement.classList.toggle("display", !!payload.DisplayMode);
        equationElement.style.fontSize = `${scale}em`;

        try {
            katex.render(payload.Expression, equationElement, {
                displayMode: !!payload.DisplayMode,
                throwOnError: true,
                output: "htmlAndMathml",
                strict: "warn",
                trust: false
            });

            return {
                ok: true,
                width: Math.ceil(equationElement.getBoundingClientRect().width),
                height: Math.ceil(equationElement.getBoundingClientRect().height)
            };
        } catch (error) {
            equationElement.innerHTML = "";

            return {
                ok: false,
                error: error instanceof Error ? error.message : String(error)
            };
        }
    }
};
