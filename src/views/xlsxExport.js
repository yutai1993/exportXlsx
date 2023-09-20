// xlsxExport.js
/**
 *纯前端 原生导出excel
 */

const trans2Base64 = (content) => {
    return window.btoa(unescape(encodeURIComponent(content)));
};

export const exportExcelFromFront = (params) => {
    const { cellList, headerList, caption, exportName = 'exportName' } = params;

    const captionEle = caption ? `<caption>${caption}</caption>` : ''; // 表格标题
    const headerEle = `<tr>${headerList?.map((item) => `<th>${item}</th>`)?.join('')}</tr>`;
    const cellEle = cellList
        ?.map((itemRow) => `<tr>${itemRow?.map((itemCell) => `<td>${itemCell}</td>`)?.join('')}</tr>`)
        ?.join('');

    const excelContent = `${captionEle}${headerEle}${cellEle}`;
    let excelFile =
        "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:x='urn:schemas-microsoft-com:office:excel' xmlns='http://www.w3.org/TR/REC-html40'>";
    excelFile +=
        '<head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head>';
    excelFile += "<body><table width='10%'  border='1'>";
    excelFile += excelContent;
    excelFile += '</table></body>';
    excelFile += '</html>';
    const link = `data:application/vnd.ms-excel;base64,${trans2Base64(excelFile)}`;
    const a = document.createElement('a');
    a.download = `${exportName}.xlsx`;
    a.href = link;
    a.click();
};
