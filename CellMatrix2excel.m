function flatMatrix = CellMatrix2excel(Matrixdata)
%cell由多个长度不同，宽度相等的矩阵构成：N*M*？， N 矩阵数量；M矩阵宽度；？矩阵长度
%将其保存到excel文件为：N*(M*?)

% 创建一个 Excel 文件名
excelFileName = 'matrix_data.xlsx';

% 创建一个 Excel 文件写入对象
excelApp = actxserver('Excel.Application');
excelApp.Visible = 1;
workbook = excelApp.Workbooks.Add;
sheet = workbook.Sheets.Item(1);

% 遍历 cell 数组中的数据，将矩阵数据按行写入工作表
for rowIndex = 1:numel(Matrixdata)
    matrix = Matrixdata{rowIndex};
 % 将矩阵扁平化为一维数组，并将其按行写入工作表
    flatMatrix = matrix(:)';
    numCols=length(flatMatrix);
 % 计算结束行和结束列的字母加数字表示,需要调用xlcolumnletter函数
    endRowLetter = xlcolumnletter(1); % 第一列的字母为 A
    endColLetter = xlcolumnletter(numCols);
% 构建范围字符串
    rangeStr = [endRowLetter num2str(rowIndex) ':' endColLetter num2str(rowIndex)]; 
     % 在工作表中创建一个大小合适的范围
    range = sheet.Range(rangeStr);
    
    % 将矩阵数据写入范围内
    range.Value = flatMatrix;

end
% 保存 Excel 文件
excelFilePath = fullfile(pwd, excelFileName);
workbook.SaveAs(excelFilePath);

% 关闭 Excel 文件
workbook.Close(false);

% 退出 Excel
excelApp.Quit;
excelApp.delete;
end