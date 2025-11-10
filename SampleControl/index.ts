import qrcode from "qrcode-generator";
import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class QRCodeCreater implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private static readonly defaultSize = 128;
    private static readonly minSize = 48;
    private static readonly containerPadding = 5;
    private static readonly placeholderText = "QRCODE";
    private hostContainer: HTMLDivElement | undefined;
    private qrWrapper: HTMLDivElement | undefined;
    private messageElement: HTMLDivElement | undefined;
    private currentValue: string | null = null;
    private currentSize: number | null = null;

    /**
     * Create the static DOM scaffolding that we reuse for every update.
     */
    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this.hostContainer = container;
        this.hostContainer.style.display = "flex";
        this.hostContainer.style.flexDirection = "column";
        this.hostContainer.style.alignItems = "center";
        this.hostContainer.style.justifyContent = "center";
        this.hostContainer.style.width = "100%";
        this.hostContainer.style.height = "100%";
        this.hostContainer.style.padding = `${QRCodeCreater.containerPadding}px`;
        this.hostContainer.style.boxSizing = "border-box";
        this.hostContainer.style.rowGap = "4px";
        this.hostContainer.style.textAlign = "center";

        this.qrWrapper = document.createElement("div");
        this.qrWrapper.style.backgroundColor = "#ffffff";
        this.qrWrapper.style.border = "1px solid #e1e1e1";
        this.qrWrapper.style.display = "none";
        this.qrWrapper.setAttribute("role", "img");
        this.qrWrapper.style.margin = "0";

        this.messageElement = document.createElement("div");
        this.messageElement.textContent = QRCodeCreater.placeholderText;
        this.messageElement.style.fontFamily = "Segoe UI, sans-serif";
        this.messageElement.style.fontSize = "12px";
        this.messageElement.style.color = "#605e5c";
        this.messageElement.style.textAlign = "center";
        this.messageElement.style.margin = "0";

        container.appendChild(this.qrWrapper);
        container.appendChild(this.messageElement);
    }

    /**
     * Re-render the QR code or placeholder whenever the bound value changes.
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void {
        if (!this.qrWrapper || !this.messageElement || !this.hostContainer) {
            return;
        }

        const rawValue = context.parameters.qrText?.raw ?? "";
        const trimmedValue = rawValue.trim();

        const qrSize = this.calculateQrSize(context);

        if (!trimmedValue) {
            this.showMessage(QRCodeCreater.placeholderText);
            this.currentValue = null;
            this.currentSize = null;
            return;
        }

        if (this.currentValue === rawValue && this.currentSize === qrSize) {
            return;
        }

        try {
            const qrInstance = qrcode(0, "M");
            qrInstance.addData(rawValue);
            qrInstance.make();

            this.qrWrapper.innerHTML = qrInstance.createSvgTag(2, 0);

            const svgElement = this.qrWrapper.querySelector("svg");
            if (svgElement) {
                svgElement.setAttribute("width", `${qrSize}`);
                svgElement.setAttribute("height", `${qrSize}`);
            }

            this.qrWrapper.style.width = `${qrSize}px`;
            this.qrWrapper.style.height = `${qrSize}px`;
            this.qrWrapper.style.display = "inline-block";
            this.qrWrapper.title = rawValue;
            this.hostContainer.title = rawValue;
            this.hostContainer.setAttribute("aria-label", `QR code for ${rawValue}`);

            this.messageElement.style.display = "none";
            this.qrWrapper.style.visibility = "visible";
            this.currentValue = rawValue;
            this.currentSize = qrSize;
        } catch (error) {
            this.showMessage("QR generation error");
            this.currentValue = null;
            this.currentSize = null;
        }
    }

    public getOutputs(): IOutputs {
        return {};
    }

    /**
     * Remove any DOM references created during init.
     */
    public destroy(): void {
        if (!this.hostContainer) {
            return;
        }

        while (this.hostContainer.firstChild) {
            this.hostContainer.removeChild(this.hostContainer.firstChild);
        }

        this.hostContainer.removeAttribute("aria-label");
        this.hostContainer = undefined;
        this.qrWrapper = undefined;
        this.messageElement = undefined;
        this.currentValue = null;
        this.currentSize = null;
    }

    private showMessage(message: string): void {
        if (!this.qrWrapper || !this.messageElement) {
            return;
        }

        this.qrWrapper.style.display = "none";
        this.qrWrapper.innerHTML = "";
        this.qrWrapper.removeAttribute("title");

        this.messageElement.textContent = message;
        this.messageElement.style.display = "block";

        if (this.hostContainer) {
            this.hostContainer.removeAttribute("title");
            this.hostContainer.removeAttribute("aria-label");
        }
    }

    private calculateQrSize(context: ComponentFramework.Context<IInputs>): number {
        const widthFromDom = this.hostContainer?.clientWidth ?? 0;
        const heightFromDom = this.hostContainer?.clientHeight ?? 0;

        const widthFromContext =
            typeof context.mode.allocatedWidth === "number" && context.mode.allocatedWidth > 0
                ? context.mode.allocatedWidth
                : 0;
        const heightFromContext =
            typeof context.mode.allocatedHeight === "number" && context.mode.allocatedHeight > 0
                ? context.mode.allocatedHeight
                : 0;

        const width = widthFromDom > 0 ? widthFromDom : widthFromContext > 0 ? widthFromContext : QRCodeCreater.defaultSize;
        const height = heightFromDom > 0 ? heightFromDom : heightFromContext > 0 ? heightFromContext : QRCodeCreater.defaultSize;

        const usable = Math.min(width, height) - QRCodeCreater.containerPadding * 2;
        return Math.max(QRCodeCreater.minSize, usable > 0 ? Math.floor(usable) : QRCodeCreater.minSize);
    }
}
