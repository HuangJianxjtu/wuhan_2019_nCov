from pyecharts import Geo, Style, Map
import xlrd
import numpy as np

# load excel sheet
excel_file = xlrd.open_workbook('./data/data.xlsx')

data_wuhan = excel_file.sheet_by_index(0)
data_mainland_china = excel_file.sheet_by_index(1)

# discard the data before 2020.01.10
infections_wuhan = np.array(data_wuhan.col_values(5, 4, data_wuhan.nrows))
total_infection_wuhan = infections_wuhan.sum()

# 省和直辖市
province_distribution = {'河南': 0, '北京': 0, '河北': 0, '辽宁': 0, '江西': 0, '上海': 0, '安徽': 0, '江苏': 0, '湖南': 0,
                         '浙江': 0, '海南': 0, '广东': 0, '湖北': 0, '黑龙江': 0, '澳门': 0, '陕西': 0, '四川': 0, '内蒙古': 0,
                         '重庆': 0, '云南': 0, '贵州': 0, '吉林': 0, '山西': 0, '山东': 0, '福建': 0, '青海': 0, '天津': 0,
                         '其他': 0}
# summary for provinces
for i_ in range(data_mainland_china.nrows-1):
    p_ = data_mainland_china.cell_value(1+i_, 2)
    if p_ in province_distribution:
        province_distribution[p_] += data_mainland_china.cell_value(1+i_, 5)
    else:
        print("没有这个省份")
province_distribution['湖北'] += total_infection_wuhan

# remove the provinces whose values're 0
dic_tmp = {}
for item_ in province_distribution.keys():
    if province_distribution[item_] != 0:
        dic_tmp[item_] = province_distribution[item_]
province_distribution = dic_tmp

provices_names = list(province_distribution.keys())
values = list(province_distribution.values())

# 数据只能是省名和直辖市的名称
map_china = Map("中国地图", '中国地图', width=1200, height=600)
map_china.add("中国地图", provices_names, values, visual_range=[0, 50], maptype='china', is_visualmap=True,
              visual_text_color='#000', type='heatmap')
map_china.show_config()
map_china.render(path="./中国地图_heatmap.html")
