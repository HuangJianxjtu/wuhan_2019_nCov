import numpy as np
from pyecharts import Geo

areas = ['广东', '广西', '湖南', '江西', '福建']
values = np.random.randint(1, 100, size=5)
test_geo = Geo("test", width=1200, height=600)
test_geo.use_theme('dark')

test_geo.add("气泡图", areas, values, maptype='china',
             is_visualmap=True,
             is_label_show=True,
             visual_type="size")  # 画地图散点十分建议指定边界颜色为一个较亮的颜色

test_geo.render(path="./气泡图.html")
