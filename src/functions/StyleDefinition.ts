export class Style {
    name: string;
    fontName: string;
    fontColor: string;
    numberFormat: string | undefined;
    fillColor: string;
    alignment: Excel.HorizontalAlignment;

    constructor(
        name: string,
        fontName: string,
        fontColor: string,
        numberFormat: string | undefined,
        fillColor: string,
        alignment: Excel.HorizontalAlignment
    ) {
        this.name = name;
        this.fontName = fontName;
        this.fontColor = fontColor;
        this.numberFormat = numberFormat;
        this.fillColor = fillColor;
        this.alignment = alignment;
    }
}

const getCollectionOfCustomtyles = () => [
    new Style('customStyleColsHeaderRow', 'Calibri', '#F2F2F2', undefined, '#262626', Excel.HorizontalAlignment.left),
    new Style(
        'customStyleColsHier1Attribute1Level1',
        'Calibri',
        '#262626',
        undefined,
        '#D9D9D9',
        Excel.HorizontalAlignment.right
    ),
    new Style(
        'customStyleColsHier1Attribute1Row',
        'Calibri',
        '#262626',
        undefined,
        '#D9D9D9',
        Excel.HorizontalAlignment.left
    ),
    new Style(
        'customStyleColsHier1ElementsRow',
        'Calibri',
        '#F2F2F2',
        undefined,
        '#404040',
        Excel.HorizontalAlignment.left
    ),
    new Style(
        'customStyleColsHier1Header',
        'Calibri',
        '#F2F2F2',
        undefined,
        '#262626',
        Excel.HorizontalAlignment.left
    ),
    new Style(
        'customStyleColsHier1HeaderAttribute1',
        'Calibri',
        '#F2F2F2',
        undefined,
        '#262626',
        Excel.HorizontalAlignment.left
    ),
    new Style(
        'customStyleColsHier1Level1',
        'Calibri',
        '#F2F2F2',
        undefined,
        '#404040',
        Excel.HorizontalAlignment.left
    ),
    new Style(
        'customStyleColsHier1Level2',
        'Calibri',
        '#F2F2F2',
        undefined,
        '#595959',
        Excel.HorizontalAlignment.left
    ),
    new Style(
        'customStyleColsHier1Level3',
        'Calibri',
        '#F2F2F2',
        undefined,
        '#808080',
        Excel.HorizontalAlignment.left
    ),
    new Style('customStyleDataValue', 'Calibri', '#595959', '#,##0.00', '#FBFBFB', Excel.HorizontalAlignment.general),
    new Style(
        'customStyleFiltersAttribute1Row',
        'Calibri',
        '#262626',
        undefined,
        '#D9D9D9',
        Excel.HorizontalAlignment.left
    ),
    new Style(
        'customStyleFiltersHeaderRow',
        'Calibri',
        '#F2F2F2',
        undefined,
        '#262626',
        Excel.HorizontalAlignment.left
    ),
    new Style(
        'customStyleFiltersHier1Attribute1',
        'Calibri',
        '#262626',
        undefined,
        '#D9D9D9',
        Excel.HorizontalAlignment.left
    ),
    new Style(
        'customStyleFiltersHier1Header',
        'Calibri',
        '#F2F2F2',
        undefined,
        '#262626',
        Excel.HorizontalAlignment.left
    ),
    new Style(
        'customStyleFiltersHier1Selection',
        'Calibri',
        '#F2F2F2',
        undefined,
        '#404040',
        Excel.HorizontalAlignment.left
    ),
    new Style(
        'customStyleFiltersSelectionRow',
        'Calibri',
        '#F2F2F2',
        undefined,
        '#404040',
        Excel.HorizontalAlignment.left
    ),
    new Style(
        'customStyleRowsHier1Attribute1Level1',
        'Calibri',
        '#262626',
        undefined,
        '#D9D9D9',
        Excel.HorizontalAlignment.left
    ),
    new Style(
        'customStyleRowsHier1Attribute1PlaceholderLevel1',
        'Calibri',
        '#D9D9D9',
        undefined,
        '#D9D9D9',
        Excel.HorizontalAlignment.left
    ),
    new Style(
        'customStyleColsHier1Attribute1PlaceholderLevel1',
        'Calibri',
        '#D9D9D9',
        undefined,
        '#D9D9D9',
        Excel.HorizontalAlignment.right
    ),
    new Style(
        'customStyleRowsHier1Header',
        'Calibri',
        '#F2F2F2',
        undefined,
        '#262626',
        Excel.HorizontalAlignment.left
    ),
    new Style(
        'customStyleRowsHier1HeaderAttribute1',
        'Calibri',
        '#F2F2F2',
        undefined,
        '#262626',
        Excel.HorizontalAlignment.left
    ),
    new Style(
        'customStyleRowsHier1Level1',
        'Calibri',
        '#F2F2F2',
        undefined,
        '#404040',
        Excel.HorizontalAlignment.left
    ),
    new Style(
        'customStyleRowsHier1Level2',
        'Calibri',
        '#F2F2F2',
        undefined,
        '#595959',
        Excel.HorizontalAlignment.left
    ),
    new Style(
        'customStyleRowsHier1Level3',
        'Calibri',
        '#F2F2F2',
        undefined,
        '#808080',
        Excel.HorizontalAlignment.left
    ),
    new Style(
        'customStyleRowsHier1PlaceholderLevel1',
        'Calibri',
        '#404040',
        undefined,
        '#404040',
        Excel.HorizontalAlignment.left
    ),
    new Style(
        'customStyleRowsHier1PlaceholderLevel2',
        'Calibri',
        '#595959',
        undefined,
        '#595959',
        Excel.HorizontalAlignment.left
    ),
    new Style(
        'customStyleRowsHier1PlaceholderLevel3',
        'Calibri',
        '#808080',
        undefined,
        '#808080',
        Excel.HorizontalAlignment.left
    ),
    new Style(
        'customStyleColsHier1PlaceholderLevel1',
        'Calibri',
        '#404040',
        undefined,
        '#404040',
        Excel.HorizontalAlignment.right
    ),
    new Style(
        'customStyleColsHier1PlaceholderLevel2',
        'Calibri',
        '#595959',
        undefined,
        '#595959',
        Excel.HorizontalAlignment.right
    ),
    new Style(
        'customStyleColsHier1PlaceholderLevel3',
        'Calibri',
        '#808080',
        undefined,
        '#808080',
        Excel.HorizontalAlignment.right
    )
];

export const addCustomStyles = (context: Excel.RequestContext) => {
    const styles = getCollectionOfCustomtyles();
    const excelStyles = context.workbook.styles;
    styles.forEach(style => {
        excelStyles.add(style.name);
        const styleItem = excelStyles.getItem(style.name);

        styleItem.font.name = style.fontName;
        styleItem.font.color = style.fontColor;
        styleItem.fill.color = style.fillColor;
        styleItem.horizontalAlignment = style.alignment;

        if (style.numberFormat !== undefined) {
            styleItem.numberFormat = style.numberFormat;
        }
    });
};
