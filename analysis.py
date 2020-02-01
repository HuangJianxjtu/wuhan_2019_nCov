from pyecharts import Geo, Style, Map
import xlrd
import numpy as np
import datetime

# load excel sheet
excel_file = xlrd.open_workbook('./data/data.xlsx')

data_wuhan = excel_file.sheet_by_index(0)
data_mainland_china = excel_file.sheet_by_index(1)

# discard the data before 2020.01.10
infections_wuhan = np.array(data_wuhan.col_values(5, 4, data_wuhan.nrows))
total_infection_wuhan = infections_wuhan.sum()

# 34个直辖市/省/自治区+港澳台
province_distribution = {'北京': 0, '天津': 0, '上海': 0, '重庆': 0, '河北': 0, '山西': 0, '辽宁': 0, '吉林': 0,
                         '黑龙江': 0, '江苏': 0, '浙江': 0, '安徽': 0, '福建': 0, '江西': 0, '山东': 0, '河南': 0,
                         '湖北': 0, '湖南': 0, '广东': 0, '海南': 0, '四川': 0, '贵州': 0, '云南': 0, '陕西': 0,
                         '甘肃': 0, '青海': 0, '内蒙古': 0, '广西': 0, '西藏': 0, '宁夏': 0, '新疆': 0,
                         '香港': 0, '澳门': 0, '台湾': 0}

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
map_china = Map("全国疫情地图-省份",
                title_color="#fff",
                title_pos="center",
                width=1200,
                height=600)
map_china.add("", provices_names, values, visual_range=[0, 100], maptype='china', is_visualmap=True,
              visual_text_color='#000', type='heatmap')
# map_china.show_config()
datetime_now = str(datetime.datetime.now())
map_china.render(path="./visual_maps/疫情地图_province_"+datetime_now+".html")

# summary for cities
city_data = {'武汉': total_infection_wuhan}
for i_ in range(data_mainland_china.nrows-1):
    city_ = data_mainland_china.cell_value(1+i_, 3)
    if city_ in city_data:
        city_data[city_] += data_mainland_china.cell_value(1+i_, 5)
    else:   # add new city
        city_data[city_] = data_mainland_china.cell_value(1 + i_, 5)

# bug
city_data.pop('广安')
city_data.pop('未知')

keys = list(city_data.keys())
values = list(city_data.values())

geo = Geo("全国疫情地图-城市",  # 设置地图标题
          title_color="#fff",
          title_pos="center",
          width=1200, height=600,
          background_color='#404a59')

geo.add("",  # 标题，构建坐标系的时候已经写好，不需要设置，设为空
        keys,
        values,
        type='effectScatter',
        visual_text_color="#fff",
        visual_range=[0, 300],
        is_visualmap=True)
# type有scatter, effectScatter, heatmap三种模式可选，可根据自己的需求选择对应的图表模式
geo.render(path="./visual_maps/疫情地图_city_"+datetime_now+".html")
