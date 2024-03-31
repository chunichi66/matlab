clear
clc
nglist={ ...
    'xy_0001_a' ...
    'xy_0002_a' ...
    'xy_0004_e' ...
    'xy_0008_d' ...
    };

checkxl("aaaa.xlsx",nglist);

function checkxl(fname,nglist)

    warcell=readcell(fname,'Range','B5:D10000');

    for i=1:10000
        warno='';
        if ismissing(warcell{i,3}) %空白行検出したら終わり
            break;
        end

        if warcell{i,3}>0 %エラー１件以上の場合
            for j=i:-1:1
               warno = warcell{j,1};
                if ~ismissing(warno) %セル合併していて空白行なら一つ上を見る
                    break;
                end
            end
            warno = strcat(warno,'_',warcell{i,2});
        end
    
        if max(contains(nglist,warno))
           error(warno);
        end
    end
end
