clear;
Tc=readcell("vbatest.xls","Sheet","仕様書","NumHeaderLines",1);

% Excel起動
excel = actxserver('Excel.Application');
excel.Visible = false;
workbook = excel.Workbooks.Add;
sheet = workbook.Sheets.Item(1);
sheet.Range('A1:A9999').NumberFormat = '@'; %文字列に設定しておく

j=1;
for i=1:height(Tc)
    dotcount=count(char(Tc(i,2)),".");
    sheet.Range(sprintf('A%d', j)).Value = Tc(i,1);

    jlast=j;
    if dotcount==0
        sheet.Range(sprintf('B%d', j)).Value = '上位要求';
        sheet.Range(sprintf('C%d:D%d', j,j)).Value = Tc(i,2:3);
        sheet.Range(sprintf('D%d:F%d', j,j)).Merge;
        sheet.Range(sprintf('G%d', j)).Value = Tc(i,6);
        j=j+1;
        sheet.Range(sprintf('C%d', j)).Value = '理由';
        sheet.Range(sprintf('D%d', j)).Value = Tc(i,4);
        sheet.Range(sprintf('D%d:F%d', j,j)).Merge;
        j=j+1;
        sheet.Range(sprintf('C%d', j)).Value = '説明';
        sheet.Range(sprintf('D%d', j)).Value = Tc(i,5);
        sheet.Range(sprintf('D%d:F%d', j,j)).Merge;

        sheet.Range(sprintf('B%d:F%d', jlast,j)).Interior.Color=255 + 200*256;
        sheet.Range(sprintf('D%d:F%d', j,j)).Merge;
        sheet.Range(sprintf('D%d:F%d', jlast,j)).WrapText = true;
    elseif dotcount==1
        sheet.Range(sprintf('C%d', j)).Value = '下位要求';
        sheet.Range(sprintf('D%d:E%d', j,j)).Value = Tc(i,2:3);
        sheet.Range(sprintf('E%d:F%d', j,j)).Merge;
        sheet.Range(sprintf('G%d', j)).Value = Tc(i,6);
        j=j+1;
        sheet.Range(sprintf('D%d', j)).Value = '理由';
        sheet.Range(sprintf('E%d:F%d', j,j)).Merge;
        sheet.Range(sprintf('E%d', j)).Value = Tc(i,4);
        j=j+1;
        sheet.Range(sprintf('D%d', j)).Value = '説明';
        sheet.Range(sprintf('E%d:F%d', j,j)).Merge;
        sheet.Range(sprintf('E%d', j)).Value = Tc(i,5);

        sheet.Range(sprintf('B%d:B%d', jlast,j)).Interior.Color=255 + 200*256;
        sheet.Range(sprintf('C%d:F%d', jlast,j)).Interior.Color=100 + 200*256 + 255*256^2;
        sheet.Range(sprintf('E%d:F%d', jlast,j)).WrapText = true;
    elseif dotcount==2
        sheet.Range(sprintf('D%d', j)).Value = '仕様';
        sheet.Range(sprintf('E%d:F%d', j,j)).Value = Tc(i,2:3);
        sheet.Range(sprintf('G%d', j)).Value = Tc(i,6);
        j=j+1;
        sheet.Range(sprintf('E%d', j)).Value = '理由';
        sheet.Range(sprintf('F%d', j)).Value = Tc(i,4);
        if ~strcmp(char(Tc(i,5)),"-") && ~strcmp(char(Tc(i,5)),"－")
            j=j+1;
            sheet.Range(sprintf('E%d', j)).Value = '説明';
            sheet.Range(sprintf('F%d', j)).Value = Tc(i,5);
        end
        sheet.Range(sprintf('B%d:B%d', jlast,j)).Interior.Color=255 + 200*256;
        sheet.Range(sprintf('C%d:C%d', jlast,j)).Interior.Color=100 + 200*256 + 255*256^2;
        sheet.Columns.Item('F').ColumnWidth = 40;
        sheet.Range(sprintf('F%d:F%d', jlast,j)).WrapText = true;
    else
        error("エラー");
    end
    j=j+1;
end

% 保存（必要に応じて）
filename = fullfile(pwd, 'example.xlsx');
workbook.SaveAs(filename);

% 終了処理
workbook.Close(false);
excel.Quit;
excel.delete;

        
