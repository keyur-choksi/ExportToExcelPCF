import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as XLSX from 'xlsx';
import "office-ui-fabric-core/dist/css/fabric.min.css";

export class ExportToExcelPCF implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private button: HTMLButtonElement;
    private iconElement: HTMLElement;
    private textElement: HTMLElement;
    private filename: string;
    private data: string;
    private sheetName: string;
    private container: HTMLDivElement;
    private tooltipText: string | null;

    // Feature toggles
    private autoWidthColumns: boolean;

    // Interaction/styling state
    private iconOnly: boolean;
    private normalBg: string;
    private normalText: string;
    private normalBorder: string;
    private hoverBg: string;
    private hoverText: string;
    private hoverBorder: string;
    private activeBg: string;
    private activeText: string;
    private focusBorder: string;
    private disabledBg: string;
    private disabledText: string;

    // Icon-specific states
    private iconNormalColor: string;
    private iconHoverBg: string;
    private iconActiveColor: string;

    constructor() { }

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this.container = container;

        this.button = document.createElement("button");
        this.button.type = 'button';
        this.button.style.border = "none";
        this.button.style.cursor = 'pointer';
        this.button.disabled = false;

        this.iconElement = document.createElement('i');
        this.textElement = document.createElement('span');
        this.button.appendChild(this.iconElement);
        this.button.appendChild(this.textElement);

        // Event handlers
        this.button.addEventListener('click', this.onButtonClick.bind(this));
        this.button.addEventListener('mouseover', () => this.applyHoverStyles());
        this.button.addEventListener('mouseout', () => this.applyNormalStyles());
        this.button.addEventListener('mousedown', () => this.applyActiveStyles());
        this.button.addEventListener('mouseup', () => this.applyHoverStyles());
        this.button.addEventListener('focus', () => {
            this.button.style.outline = 'none';
            if (!this.iconOnly) {
                this.button.style.borderColor = this.focusBorder;
            }
        });
        this.button.addEventListener('blur', () => this.applyNormalStyles());

        // Initialize from parameters
        this.updateButtonStyles(context.parameters);
        this.sheetName = context.parameters.ExportSheetName.raw ?? 'Exported_Data';
        this.autoWidthColumns = context.parameters.AutoWidthColumns.raw ?? false;

        this.tooltipText = context.parameters.ToolTip.raw;
        if (this.tooltipText) {
            this.container.setAttribute("title", this.tooltipText);
        }

        container.id = "export-to-excel-container";
        container.appendChild(this.button);
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this.data = context.parameters.DataToExport.raw ?? this.data;
        this.filename = (context.parameters.ExportFileName.raw ?? "Exported_Excel_File") + ".xlsx";
        this.sheetName = context.parameters.ExportSheetName.raw ?? this.sheetName;
        this.autoWidthColumns = context.parameters.AutoWidthColumns.raw ?? false;

        this.updateButtonStyles(context.parameters);

        const newTip = context.parameters.ToolTip.raw;
        if (newTip !== this.tooltipText) {
            this.tooltipText = newTip;
            if (newTip) {
                this.container.setAttribute("title", newTip);
            } else {
                this.container.removeAttribute("title");
            }
        }
    }

    private updateButtonStyles(params: IInputs): void {
        // Determine icon vs button mode
        const showIcon = params.ShowIcon.raw ?? false;
        this.iconOnly = showIcon;

        // Map enum to icon key
        const iconOption = params.IconName.raw;
        const iconName = iconOption === "1" ? "ExcelDocument" : "Download";

        // Icon colors and sizing
        this.iconNormalColor = params.IconColor.raw ?? 'RGBA(0, 120, 212, 1)';
        this.iconActiveColor = params.ActiveTextColor.raw ?? this.iconNormalColor;
        this.iconHoverBg = 'RGBA(0, 0, 0, 0.05)';

        this.iconElement.className = `ms-Icon ms-Icon--${iconName}`;
        this.iconElement.style.color = this.iconNormalColor;
        this.iconElement.style.fontSize = params.IconSize.raw ?? '14px';
        this.iconElement.style.display  = 'inline-block';

        if (this.iconOnly) {
            // Icon-only mode
            this.textElement.style.display = 'none';
            this.iconElement.style.margin = '0';
            this.button.style.backgroundColor = 'transparent';
            this.button.style.border = 'none';
            this.button.style.padding = '6px';
            this.button.style.width = 'auto';
            this.button.style.height = 'auto';
        } else {
            // Full-button mode
            this.textElement.style.display = 'inline';
            this.textElement.textContent = params.ButtonText.raw ?? 'Export to Excel';
            this.iconElement.style.marginRight = showIcon ? '4px' : '0';
            this.iconElement.style.display = showIcon ? 'inline-block' : 'none';

            // Spacing
            this.button.style.padding = params.Padding.raw ?? '8px 16px';
            this.button.style.width = params.ButtonWidth.raw ?? 'auto';
            this.button.style.height = params.ButtonHeight.raw ?? 'auto';
            this.button.style.margin = params.Margin.raw ?? '0px';

            // Typography
            this.button.style.fontFamily = params.ButtonFont.raw ?? 'Segoe UI, sans-serif';
            this.button.style.fontSize = params.ButtonTextSize.raw ?? '14px';
            // Map enum index to actual CSS values
            const fontWeightOptions = ['100','200','300','400','500','600','700','800','900'];
            const fwIndex = parseInt(params.FontWeight.raw ?? '5', 10);
            this.button.style.fontWeight = fontWeightOptions[fwIndex] || '600';
            const fontStyleOptions = ['normal','italic','oblique'];
            const fsIndex = parseInt(params.FontStyle.raw ?? '0', 10);
            this.button.style.fontStyle = fontStyleOptions[fsIndex] || 'normal';
            const textDecorationOptions = ['none','underline','overline','line-through','blink'];
            const tdIndex = parseInt(params.TextDecoration.raw ?? '0', 10);
            this.button.style.textDecoration = textDecorationOptions[tdIndex] || 'none';

            // Border & radius
            const borderCol = params.BorderColor.raw ?? 'transparent';
            this.button.style.borderColor = borderCol;
            this.button.style.borderWidth = params.BorderWidth.raw ?? '1px';
            const borderStyleOptions = ['none','hidden','dotted','dashed','solid','double','groove','ridge','inset','outset'];
            const bsIndex = parseInt(params.BorderStyle.raw ?? '4', 10);
            this.button.style.borderStyle = borderStyleOptions[bsIndex] || 'solid';
            this.button.style.borderRadius = params.ButtonRadius.raw ?? '4px';

            // Color states
            this.normalBg = params.ButtonBackgroundColor.raw ?? 'RGBA(0, 120, 212, 1)';
            this.normalText = params.ButtonTextColor.raw ?? 'RGBA(255, 255, 255, 1)';
            this.normalBorder = borderCol;
            this.hoverBg = params.HoverBackgroundColor.raw ?? 'RGBA(16, 110, 190, 1)';
            this.hoverText = params.HoverTextColor.raw ?? this.normalText;
            this.hoverBorder = params.HoverBorderColor.raw ?? this.normalBorder;
            this.activeBg = params.ActiveBackgroundColor.raw ?? 'RGBA(0, 90, 160, 1)';
            this.activeText = params.ActiveTextColor.raw ?? this.normalText;
            this.focusBorder = params.FocusBorderColor.raw ?? this.normalBorder;
            this.disabledBg = params.DisabledBackgroundColor.raw ?? 'RGBA(243, 242, 241, 1)';
            this.disabledText = params.DisabledTextColor.raw ?? this.normalText;

            this.applyNormalStyles();
        }
    }

    private applyNormalStyles(): void {
        if (this.iconOnly) {
            this.button.style.backgroundColor = 'transparent';
            this.iconElement.style.color = this.iconNormalColor;
        } else {
            this.button.style.backgroundColor = this.normalBg;
            this.button.style.color = this.normalText;
            this.textElement.style.color   = this.normalText; 
            this.button.style.borderColor = this.normalBorder;
        }
    }

    private applyHoverStyles(): void {
        if (this.iconOnly) {
            this.button.style.backgroundColor = this.iconHoverBg;
            this.iconElement.style.color = this.iconNormalColor;
        } else {
            this.button.style.backgroundColor = this.hoverBg;
            this.button.style.color = this.hoverText;
            this.textElement.style.color = this.hoverText;
            this.button.style.borderColor = this.hoverBorder;
        }
    }

    private applyActiveStyles(): void {
        if (this.iconOnly) {
            this.iconElement.style.color = this.iconActiveColor;
        } else {
            this.button.style.backgroundColor = this.activeBg;
            this.button.style.color = this.activeText;
            this.textElement.style.color      = this.activeText;
        }
    }

    private onButtonClick(event: Event): void {
        if (!this.data) return;
        let dataArray: any[];
        try {
            dataArray = JSON.parse(this.data);
        } catch (e) {
            console.warn('ExportToExcelPCF: Invalid JSON data', e);
            return;
        }

        const worksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(dataArray);

        // Apply auto-width if enabled
        if (this.autoWidthColumns && dataArray.length > 0) {
            const headers = Object.keys(dataArray[0]);
            const colWidths = headers.map(hdr => {
                const maxLen = dataArray.reduce((max, row) => {
                    const cell = row[hdr] != null ? String(row[hdr]) : '';
                    return Math.max(max, cell.length);
                }, hdr.length);
                return { wch: maxLen + 2 };
            });
            worksheet['!cols'] = colWidths;
        }

        const workbook: XLSX.WorkBook = { Sheets: { [this.sheetName]: worksheet }, SheetNames: [this.sheetName] };
        XLSX.writeFile(workbook, this.filename);
    }

    public getOutputs(): IOutputs {
        return {};
    }

    public destroy(): void { }
}