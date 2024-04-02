import os
import re
import sys
import shutil
import subprocess
import platform
import time
from datetime import datetime
import logging
from logging.handlers import TimedRotatingFileHandler
import pandas as pd
from pyecharts.charts import Bar, Line, Page, Pie
from pyecharts import options as opts
from pyecharts.commons.utils import JsCode
from pyecharts.components import Table
from pyecharts.render import make_snapshot
from snapshot_selenium import snapshot

import docxtpl
from docx.shared import Mm

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from collections import Counter

# 获取当前时间并格式化
current_time = datetime.now().strftime("%Y-%m-%d_%H_%M_%S")

BASE_PATH = os.path.split(os.path.realpath(__file__))[0]
if getattr(sys, 'frozen', False):
    BASE_PATH = os.path.dirname(sys.executable)

log_file = os.path.join(BASE_PATH, 'rpa.log')
# 设置日志的输出格式
logger = logging.getLogger()
logger.setLevel(logging.INFO)

# 创建一个文件handler
file_handler = TimedRotatingFileHandler(log_file, when="midnight", interval=1, backupCount=7, encoding='utf-8')
file_handler.setFormatter(
    logging.Formatter('%(asctime)s - %(name)s - %(levelname)s: %(message)s', datefmt='%Y-%m-%d %H:%M:%S'))
logger.addHandler(file_handler)

# 创建一个控制台handler
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.DEBUG)
console_handler.setFormatter(logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - [%(filename)s:%('
                                               'lineno)d] %(message)s', datefmt='%Y-%m-%d %H:%M:%S'))
logger.addHandler(console_handler)

def make_pie1(excel_file_path, detail_sheet_name,project_info_sheet_name):
    logger.info('正在按明细帐生成区域合同金额饼图')
    # 读取项目信息表，赋值到project_info_df
    project_info_df = pd.read_excel(excel_file_path, sheet_name=project_info_sheet_name, header=[0],
                                    engine='openpyxl')
    # 删除指定列'区内/区外'为空的数据
    project_info_df.dropna(subset=['区内/区外'], inplace=True)
    # 读取明细帐表，赋值到detail_df
    detail_df = pd.read_excel(excel_file_path, sheet_name=detail_sheet_name, skiprows=2, header=[0,1],
                                    engine='openpyxl')
    # 使用dropna方法删除'核算维度'列为空的数据
    detail_df.dropna(subset=[('核算维度', 'Unnamed: 3_level_1')], inplace=True)
    # 获取所有核算维度列的项目信息，使用.values.tolist()方法转成列表，存到project_info_list里
    project_info_list = detail_df[('核算维度', 'Unnamed: 3_level_1')].values.tolist()
    # 遍历project_info_list所有的核算维度项目信息，item就是每个项目信息
    for item in project_info_list:
        # 用正则表达式匹配项目信息列表里的每条数据，获取项目编号，存到新的列表里
        # 用正则表达式匹配项目ID，表示匹配以“项目:”开头的，由数字和字母组成的一段字符
        pattern = r"(?<=项目:)([A-Za-z0-9-]+)"
        # 在项目信息item中匹配项目ID，将结果赋值到match
        match = re.search(pattern, item)
        # 如果match不为空，就是匹配到了
        if match:
            # 就获取项目id赋值到project_id
            priject_id = match.group(0)
            # 将匹配到的项目ID添加到到明细帐表的 项目信息->项目编号 列里
            detail_df.loc[detail_df[('核算维度', 'Unnamed: 3_level_1')] == item, ('项目信息', '项目编号')] = priject_id
            # 用获取到的项目id从项目信息表匹配项目区域信息，得到这个项目是区内还是区外的信息，赋值到area
            area = project_info_df.loc[project_info_df['项目编号'] == priject_id,'区内/区外']
            # 处理异常，如果匹配到项目区域，就加入‘区域’列里，匹配不到就不做任何处理
            try:
                # 匹配到区内区外信息，存到明细帐表的项目信息->区域列中
                detail_df.loc[detail_df[('核算维度', 'Unnamed: 3_level_1')] == item, ('项目信息', '区域')] = area.values[0]
            except IndexError:
                pass
    # 使用dropna方法删除'项目编号'列为空的数据
    detail_df.dropna(subset=[('项目信息', '项目编号')], inplace=True)
    # 使用dropna方法删除'区域'列为空的数据
    detail_df.dropna(subset=[('项目信息', '区域')], inplace=True)
    # 按照'区内/区外'列分组，计算每个分组的'本期发生额','贷方金额'之和
    inside_the_area = detail_df.loc[detail_df[('项目信息', '区域')] == '区内', ('本期发生额','贷方金额')].sum()
    outside_the_area = detail_df.loc[detail_df[('项目信息', '区域')] == '区外',('本期发生额','贷方金额')].sum()

    # 如果数据为空，就不执行下面的操作直接返回
    if not inside_the_area and not outside_the_area:
        return None
    # 创建一个饼图对象
    c = (
        Pie()
        # 添加数据，将区域内的数据和区域外的数据分别作为第一个参数和第二个参数传入
        .add("", [list(z) for z in zip(['区内', '区外'], [inside_the_area, outside_the_area])])
        # 设置图例选项，隐藏图例
        .set_global_opts(legend_opts=opts.LegendOpts(is_show=False))
        # 设置系列选项，设置标签的格式和字体大小
        .set_series_opts(label_opts=opts.LabelOpts(formatter="{b}: {d}%", font_size=20))
    )
    # 渲染饼图，并将结果保存为html文件
    c.render("按区域分析饼图.html")
    # 获取当前时间，并设置图片的路径
    current_time = datetime.now().strftime("%Y-%m-%d_%H_%M_%S")
    image_path = os.path.join(BASE_PATH, 'images', f'area_pie1_{current_time}.png')
    # 将饼图保存为图片文件
    make_snapshot_to_file(c, image_path)
    # 返回图片路径
    return image_path

def make_pie6(excel_file_path, sheet_name):
    logger.info('正在按区域分析生成合同金额饼图')
    # 生成饼图
    project_info_df = pd.read_excel(excel_file_path, sheet_name=sheet_name, header=[0],
                                    engine='openpyxl')
    # 删除指定列'区内/区外'为空的数据
    project_info_df.dropna(subset=['区内/区外'], inplace=True)
    # 按照'区内/区外'列分组，计算每个分组的合同金额之和
    grouped = project_info_df.groupby('区内/区外')['合同金额(元)'].sum()
    # 获取'区内'的合同金额
    inside_the_area = float(grouped.get('区内', 0))
    # 获取'区外'的合同金额
    outside_the_area = float(grouped.get('区外', 0))
    # 如果数据为空，就不执行下面的操作直接返回
    if not inside_the_area and not outside_the_area:
        return None
    # 创建一个饼图对象
    c = (
        Pie()
        # 添加数据，将区域内的数据和区域外的数据分别作为第一个参数和第二个参数传入
        .add("", [list(z) for z in zip(['区内', '区外'], [inside_the_area, outside_the_area])])
        # 设置图例选项，隐藏图例
        .set_global_opts(legend_opts=opts.LegendOpts(is_show=False))
        # 设置系列选项，设置标签的格式和字体大小
        .set_series_opts(label_opts=opts.LabelOpts(formatter="{b}: {d}%", font_size=20))
    )
    # 渲染饼图，并将结果保存为html文件
    c.render("按区域分析饼图.html")
    # 获取当前时间，并设置图片的路径
    current_time = datetime.now().strftime("%Y-%m-%d_%H_%M_%S")
    image_path = os.path.join(BASE_PATH, 'images', f'area_pie1_{current_time}.png')
    # 将饼图保存为图片文件
    make_snapshot_to_file(c, image_path)
    # 返回图片路径
    return image_path


def make_pie2(excel_file_path, sheet_name):
    logger.info('正在按区内合同金额生成饼图')
    # 生成饼图
    # 读取指定路径的excel文件，并将其存储为dataframe
    project_info_df = pd.read_excel(excel_file_path, sheet_name=sheet_name, header=[0],
                                        engine='openpyxl')
    # 删除空值
    project_info_df.dropna(subset=['区内/区外'], inplace=True)
    # 筛选出区内数据
    inside_the_area = project_info_df[project_info_df['区内/区外'] == '区内']
    # 对数据进行分组，计算每个部门全称下所有合同金额的总和
    grouped = inside_the_area.groupby(['区内/区外', '部门全称'])['合同金额(元)'].sum()
        # print(grouped.keys(),grouped.values)

    # 将分组结果存储到字典中
    projects = grouped.to_dict()
    new_dic = {}
    for item in projects.keys():
        new_dic[item[1]] = projects[item]
    # 如果字典为空，就不执行后续生成图表的操作
    if not new_dic:
        return None
    c = (
        Pie()
        .add("", [list(z) for z in zip(new_dic.keys(), new_dic.values())])
        .set_global_opts(legend_opts=opts.LegendOpts(is_show=False))
        .set_series_opts(label_opts=opts.LabelOpts(formatter="{b}: {d}%", font_size=20))
    )
    c.render("按区内项目类别分析饼图.html")
    # 获取当前时间并格式化，用于拼接图表文件名
    current_time = datetime.now().strftime("%Y-%m-%d_%H_%M_%S")
    # 拼接图表文件名
    image_path = os.path.join(BASE_PATH, 'images', f'area_pie2_{current_time}.png')
    # 生成图表
    make_snapshot_to_file(c, image_path)
    # 返回生成的图表路径
    return image_path


def make_pie3(excel_file_path, sheet_name):
    logger.info('正在按区外合同金额分析生成饼图')
    # 生成饼图
    # 读取指定路径的excel文件，指定工作表名称，并设置表头为第一行
    project_info_df = pd.read_excel(excel_file_path, sheet_name=sheet_name, header=[0],
                                    engine='openpyxl')
    # 清除空值
    project_info_df.dropna(subset=['区内/区外'], inplace=True)
    # 提取区外数据
    outside_the_area = project_info_df[project_info_df['区内/区外'] == '区外']
    # 按指定字段分组，并计算合同金额总和
    grouped = outside_the_area.groupby(['区内/区外', '部门全称'])['合同金额(元)'].sum()

    # 将分组结果存储到字典中
    # 将grouped数据转换为字典
    projects = grouped.to_dict()
    # 创建一个新的字典
    new_dic = {}
    # 遍历projects字典的键
    for item in projects.keys():
        # 将键的元组中的第二个元素作为新字典的键
        new_dic[item[1]] = projects[item]

    # 如果不存在新字典，则返回None
    if not new_dic:
        return None

    # 创建一个饼图对象
    c = (
        Pie()
        # 添加数据，将new_dic中的键值对分别作为参数传递给Pie的add方法
        .add("", [list(z) for z in zip(new_dic.keys(), new_dic.values())])
        # 设置图例选项，隐藏图例
        .set_global_opts(legend_opts=opts.LegendOpts(is_show=False))
        # 设置系列选项，设置标签的格式和字体大小
        .set_series_opts(label_opts=opts.LabelOpts(formatter="{b}: {d}%", font_size=20))
    )
    # 渲染饼图，并将结果保存为html文件
    c.render("按区外项目类别分析饼图.html")
    # 获取当前时间，并设置图片的路径
    current_time = datetime.now().strftime("%Y-%m-%d_%H_%M_%S")
    image_path = os.path.join(BASE_PATH, 'images', f'area_pie3_{current_time}.png')
    # 将饼图保存为图片文件
    make_snapshot_to_file(c, image_path)
    # 返回图片路径
    return image_path

def make_pie4(excel_file_path, sheet_name):
    logger.info('正在按区外项目类别分析生成饼图')

    # 读取Excel工作簿数据, 跳过前两行的说明性文本,二级表头
    summary_2023 = pd.read_excel(excel_file_path, sheet_name=sheet_name, skiprows=2, header=[0, 1],
                                 engine='openpyxl')
    # 去除所有列中的前后空格
    summary_2023 = summary_2023.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    # 剔除核算组织是空的行
    summary_2023.dropna(subset=[('核算组织', 'Unnamed: 2_level_1')], inplace=True)
    summary_2023 = summary_2023[~summary_2023[('核算组织', 'Unnamed: 2_level_1')].str.contains('集团')]
    # 将金额单位从元改为万元
    summary_2023[('本期发生额', '贷方金额')] = round(summary_2023[('本期发生额', '贷方金额')] / 10000,2)
    # 获取2023年的核算组织列表
    department_2023 = summary_2023[('核算组织', 'Unnamed: 2_level_1')].values.tolist()

    # 获取当前期期初余额
    current_amount = get_df_data(summary_2023, department_2023, ('核算组织', 'Unnamed: 2_level_1'),
                                 ('本期发生额', '贷方金额'))

    if not current_amount:
        return None

    # 创建饼图
    c = (
        Pie()
        .add("", [list(z) for z in zip(current_amount.keys(), current_amount.values())])
        .set_global_opts(legend_opts=opts.LegendOpts(is_show=False))
        .set_series_opts(label_opts=opts.LabelOpts(formatter="{b}: {d}%", font_size=20))
    )
    # 渲染饼图
    c.render("按项目类别分析饼图.html")
    # 获取当前时间
    current_time = datetime.now().strftime("%Y-%m-%d_%H_%M_%S")
    # 拼接图片路径
    image_path = os.path.join(BASE_PATH, 'images', f'area_pie3_{current_time}.png')
    # 将饼图保存到文件
    make_snapshot_to_file(c, image_path)
    # 返回图片路径
    return image_path

def make_pie5(excel_file_path, sheet_name):
    logger.info('正在按区外项目类别分析生成饼图')
    # 生成饼图
    # 读取Excel文件，指定工作表名称和工作表名称
    project_info_df = pd.read_excel(excel_file_path, sheet_name=sheet_name, header=[0],
                                    engine='openpyxl')
    # 删除工作表中的空值
    project_info_df.dropna(subset=['区内/区外'], inplace=True)
    # 对工作表中的合同金额进行分组，并保留万元为单位
    grouped = project_info_df.groupby(['部门全称'])['合同金额(元)'].sum() / 10000

    # 将分组结果存储到字典中
    projects = grouped.to_dict()

    if not projects:
        return None

    # 创建饼图
    c = (
        Pie()
        .add("", [list(z) for z in zip(projects.keys(), projects.values())])
        .set_global_opts(legend_opts=opts.LegendOpts(is_show=False))
        .set_series_opts(label_opts=opts.LabelOpts(formatter="{b}: {d}%", font_size=20))
    )
    # 渲染饼图
    c.render("按项目类别分析饼图.html")
    # 获取当前时间
    current_time = datetime.now().strftime("%Y-%m-%d_%H_%M_%S")
    # 拼接图片路径
    image_path = os.path.join(BASE_PATH, 'images', f'area_pie3_{current_time}.png')
    # 将饼图保存到文件
    make_snapshot_to_file(c, image_path)
    # 返回图片路径
    return image_path

def get_df_data(df, keys, index1, index2):
    # 从按组织汇总 表格里获取数据
    res = {}
    for key in keys:
        try:
            v = df.loc[
                df[index1] == key, index2
            ].values[0]
        except:
            v = 0

        res[key] = float(v)
    return res

def make_bar1(excel_file_path, current_year_sheet_name, last_year_sheet_name, targets_sheet_name):
    logging.info('正在生成柱状图')
    # 读取Excel工作簿数据, 跳过前两行的说明性文本,二级表头
    summary_2022 = pd.read_excel(excel_file_path, sheet_name=last_year_sheet_name, skiprows=2, header=[0, 1],
                                 engine='openpyxl')
    # 去除所有列中的前后空格
    summary_2022 = summary_2022.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    # 剔除核算组织是空的行
    summary_2022.dropna(subset=[('核算组织', 'Unnamed: 2_level_1')], inplace=True)
    # 把本期的发生额金额除以10000，单位从元改为万元
    summary_2022[('本期发生额', '借方金额')] = summary_2022[('本期发生额', '借方金额')] / 10000
    # 删除集团组织所在的行
    summary_2022 = summary_2022[~summary_2022[('核算组织', 'Unnamed: 2_level_1')].str.contains('集团')]

    # 读取Excel工作簿数据, 跳过前两行的说明性文本,二级表头
    summary_2023 = pd.read_excel(excel_file_path, sheet_name=current_year_sheet_name, skiprows=2, header=[0, 1],
                                 engine='openpyxl')
    # 去除所有列中的前后空格
    summary_2023 = summary_2023.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    # 剔除核算组织是空的行
    summary_2023.dropna(subset=[('核算组织', 'Unnamed: 2_level_1')], inplace=True)
    summary_2023 = summary_2023[~summary_2023[('核算组织', 'Unnamed: 2_level_1')].str.contains('集团')]
    # 将金额单位从元改为万元
    summary_2023[('本期发生额', '借方金额')] = summary_2023[('本期发生额', '借方金额')] / 10000
    summary_2023[('本年累计', '借方金额')] = summary_2023[('本年累计', '借方金额')] / 10000

    # 获取度经营业绩考核目标
    targets_2023 = pd.read_excel(excel_file_path, sheet_name=targets_sheet_name, skiprows=1, engine='openpyxl')
    # 去除所有列中的前后空格
    targets_2023 = targets_2023.applymap(lambda x: x.strip() if isinstance(x, str) else x)


    # 获取所有核算组织
    # 获取2023年的核算组织列表
    department_2023 = summary_2023[('核算组织', 'Unnamed: 2_level_1')].values.tolist()
    # 获取2022年的核算组织列表
    department_2022 = summary_2022[('核算组织', 'Unnamed: 2_level_1')].values.tolist()
    # department_2022和department_2023列表相元素加并去重
    department_all = list(set(department_2023 + department_2022))

    sort = {
        'ZN事业部': 1,
        'ZX事业部': 2,
        'KC事业部': 3,
        'SZ设计院': 4,
        'BR公司': 5,
        'GC公司': 6,
        'J2公司': 7,
        'SZ公司': 8,
        'XCL公司': 9
    }
    # 给组织进行排序
    department_all = sorted(department_all, key=lambda x: sort.get(x, float('inf')))

    # 2023年数据
    # 获取当前期期初余额
    current_amount = get_df_data(summary_2023, department_all, ('核算组织', 'Unnamed: 2_level_1'), ('本期发生额', '借方金额'))
    # 获取当前期末余额
    cumulative_amount = get_df_data(summary_2023, department_all, ('核算组织', 'Unnamed: 2_level_1'), ('本年累计', '借方金额'))
    # 2022年数据本期借方金额用户计算同比
    last_years_amount = get_df_data(summary_2022, department_all, ('核算组织', 'Unnamed: 2_level_1'), ('本期发生额', '借方金额'))
    # 计算同比数据
    # 定义一个字典，用于存储年增长率
    Year_on_year_growth_rate_dic = {}
    # 遍历current_amount字典中的每一个key
    for key in current_amount.keys():
        # 如果last_years_amount字典中对应的value为0
        if last_years_amount[key] == 0:
            # 将年增长率设置为100
            Year_on_year_growth_rate_dic[key] = 100
        else:
            # 将年增长率计算并设置为last_years_amount字典中对应的value
            Year_on_year_growth_rate_dic[key] = round(
                ((current_amount[key] - last_years_amount[key]) / last_years_amount[key]) * 100, 2)

    targets_data = get_df_data(targets_2023, department_all, '组织', '营业收入')
    # 目标收入柱状为了避免直接堆叠，需要减去本年累计收入和本期累计收入
    target_difference = {key: targets_data[key] - cumulative_amount.get(key, 0) - current_amount.get(key,0) for key in targets_data}
    target_difference = {key: target_difference[key] if target_difference[key] > 0 else 0 for key in target_difference}

    # 计算累计/目标的百分比
    percentages = {key: round((cumulative_amount[key] / targets_data[key] * 100), 2) if targets_data[key] else 0 for key
                   in cumulative_amount}
    # 创建一个柱状图对象
    bar = Bar()
    # 添加部门做横坐标
    bar.add_xaxis(list(current_amount.keys()))
    # 添加部门数据做纵坐标
    bar.add_yaxis("本期收入", list(current_amount.values()), stack="stack1", category_gap="50%")
    bar.add_yaxis("本年累计收入", list(cumulative_amount.values()), stack="stack1", category_gap="50%")
    bar.add_yaxis("本年目标剩余量", list(target_difference.values()), stack="stack1", category_gap="50%")
    # 设置全局配置项
    bar.set_global_opts(
        legend_opts=opts.LegendOpts(
            pos_top="5%",  # 图例距离顶部5%位置
        ),
        xaxis_opts=opts.AxisOpts(
            # name="单位部门",
            axislabel_opts=opts.LabelOpts(rotate=-15)
        ),
        yaxis_opts=opts.AxisOpts(name="单位：万元")

    )

    bar.extend_axis(
        yaxis=opts.AxisOpts(
            type_="value",
            axislabel_opts=opts.LabelOpts(formatter="{value} %"),
        )
    )

    # 设置系列配置项，调整标签位置
    bar.set_series_opts(
        label_opts=opts.LabelOpts(position="top", is_show=False),
    )

    line = Line()
    # 部门做横坐标
    line.add_xaxis(list(current_amount.keys()))
    # 部门数据做纵坐标
    line.add_yaxis("同比", list(Year_on_year_growth_rate_dic.values()), yaxis_index=1)

    line.set_series_opts(
        label_opts=opts.LabelOpts(position="top", formatter=JsCode("""
                function(x) {
                    return '同比:' + x.value[1] + '%';
                }
            """),)
    )

    # 创建一个Line对象
    line1 = Line()
    # 部门做横坐标
    line1.add_xaxis(list(current_amount.keys()))
    # 添加目标series，并设置y轴索引为1
    line1.add_yaxis("目标", list(percentages.values()), yaxis_index=1, )

    # 设置series的标签
    line1.set_series_opts(
        label_opts=opts.LabelOpts(position="top", formatter=JsCode("""
                function(x) {
                    return '目标:' + x.value[1] + '%';
                }
            """), padding=20)
    )

    # 创建一个表格
    table = Table()
    # 将目标、本期、累计、完成率、增长率等数据添加到表格中
    target_amount_list = ["目标"] + list(targets_data.values())
    current_period_list = ["本期"] + list(current_amount.values())
    accumulative_total_list = ["累计"] + list(cumulative_amount.values())
    Target_completion_rate_list = ["目标完成率"] + list(percentages.values())
    Year_on_year_growth_rate_list = ["同比增长率"] + list(Year_on_year_growth_rate_dic.values())

    # 设置表头为类别
    headers = ['类别'] + list(current_amount.keys())
    # 向表格中添加数据
    rows = [
        target_amount_list,
        Target_completion_rate_list,
        current_period_list,
        accumulative_total_list,
        Year_on_year_growth_rate_list,
    ]
    table.add(headers, rows)

    # 输出图表和表格
    page = Page()
    page.add(
        bar.overlap(line).set_series_opts(z=5),
        table,
    )
    page.render("按组织分析柱状图.html")
    image_path = make_snapshot_to_file(page)
    # 返回生成的图表路径
    return image_path


def make_snapshot_to_file(c, image_path=''):
    # 如果图片路径为空，则使用当前时间作为图片名
    if not image_path:
        current_time = datetime.now().strftime("%Y-%m-%d_%H_%M_%S")
        image_path = os.path.join(BASE_PATH, 'images', f'snapshot_{current_time}.png')
    # 检查文件夹是否存在
    if not os.path.exists(os.path.join(BASE_PATH, 'images')):
        # 如果不存在，则创建文件夹
        os.makedirs(os.path.join(BASE_PATH, 'images'))
    # 渲染图表到png文件
    make_snapshot(snapshot, c.render(), image_path)
    # 删除生成的文件
    os.remove('render.html')
    return image_path


def make_doc(word_dic,image_dic, output_docx=None):
    # 输出文件路径
    if not output_docx:
        output_docx = os.path.join(BASE_PATH, "营业收入分析.docx")
    # 要编辑的docx文档模板路径
    template_doc = os.path.join(BASE_PATH, "template.docx")
    # 创建docx对象
    daily_docx = docxtpl.DocxTemplate(template_doc)
    # 创建图片对象
    context = {}
    for key in image_dic.keys():
        if not image_dic[key]:
            continue
        # 如果是pie4和pie1饼状图表，则设置宽度为140mm
        if key in ['pie4','pie1']:
            image = docxtpl.InlineImage(daily_docx, image_dic[key], width=Mm(140))
        else:
            # 否则设置为180mm
            image = docxtpl.InlineImage(daily_docx, image_dic[key], width=Mm(180))
        context[key] = image
    # 拼接图表字典和文字字典
    for key in word_dic.keys():
        context[key] = word_dic[key]

    # 渲染docx
    daily_docx.render(context)
    # 保存docx
    daily_docx.save(output_docx)
    return output_docx


def copy_file(source_file, target_file):
    try:
        # 复制源文件到目标文件
        shutil.copy2(source_file, target_file)
        return True
    except:
        return False

def make_paragraph1(excel_file_path, current_year_sheet_name, last_year_sheet_name, targets_sheet_name):
    # 读取Excel工作簿数据, 跳过前两行的说明性文本,二级表头
    summary_2022 = pd.read_excel(excel_file_path, sheet_name=last_year_sheet_name, skiprows=2, header=[0, 1],engine='openpyxl')
    # 去除所有列中的前后空格
    summary_2022 = summary_2022.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    # 剔除核算组织是空的行
    summary_2022.dropna(subset=[('核算组织', 'Unnamed: 2_level_1')], inplace=True)
    # 将金额单位从元改为万元
    summary_2022[('本期发生额', '借方金额')] = round(summary_2022[('本期发生额', '借方金额')] / 10000,2)
    # 剔除包含'集团'的组织
    summary_2022 = summary_2022[~summary_2022[('核算组织', 'Unnamed: 2_level_1')].str.contains('集团')]
    # 读取Excel工作簿数据, 跳过前两行的说明性文本,二级表头
    summary_2023 = pd.read_excel(excel_file_path, sheet_name=current_year_sheet_name, skiprows=2, header=[0, 1],engine='openpyxl')
    # 去除所有列中的前后空格
    summary_2023 = summary_2023.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    # 剔除核算组织是空的行
    summary_2023.dropna(subset=[('核算组织', 'Unnamed: 2_level_1')], inplace=True)
    # 剔除包含'集团'的组织
    summary_2023 = summary_2023[~summary_2023[('核算组织', 'Unnamed: 2_level_1')].str.contains('集团')]
    # 将金额单位从元改为万元
    summary_2023[('本期发生额', '借方金额')] = round(summary_2023[('本期发生额', '借方金额')] / 10000,2)
    summary_2023[('本年累计', '借方金额')] = round(summary_2023[('本年累计', '借方金额')] / 10000,2)

    # 获取度经营业绩考核目标
    targets_2023 = pd.read_excel(excel_file_path, sheet_name=targets_sheet_name, skiprows=1, engine='openpyxl')
    # 去除所有列中的前后空格
    targets_2023 = targets_2023.applymap(lambda x: x.strip() if isinstance(x, str) else x)

    # 获取所有核算组织
    department_2023 = summary_2023[('核算组织', 'Unnamed: 2_level_1')].values.tolist()
    department_2022 = summary_2022[('核算组织', 'Unnamed: 2_level_1')].values.tolist()
    # department_2022和department_2023列表相元素加并去重
    department_all = list(set(department_2023 + department_2022))

    # 2023年数据
    # 本期借方金额 current_amount
    current_amount = get_df_data(summary_2023, department_all, ('核算组织', 'Unnamed: 2_level_1'),
                                 ('本期发生额', '借方金额'))
    # 本期累计金额 cumulative_amount
    cumulative_amount = get_df_data(summary_2023, department_all, ('核算组织', 'Unnamed: 2_level_1'),
                                    ('本年累计', '借方金额'))


    # 2022年数据本期借方金额用户计算同比
    last_years_amount = get_df_data(summary_2022, department_all, ('核算组织', 'Unnamed: 2_level_1'),
                                    ('本期发生额', '借方金额'))
    # 计算同比数据
    # 定义一个字典，用于存储年增长率
    Year_on_year_growth_rate_dic = {}
    # 遍历current_amount字典中的每一个key
    for key in current_amount.keys():
        # 如果last_years_amount字典中key对应的值等于0
        if last_years_amount[key] == 0:
            # 将key对应的值设置为100
            Year_on_year_growth_rate_dic[key] = 100
        else:
            # 将key对应的值设置为当前值减去上一年的值，再除以上一年的值，最后乘以100，保留两位小数
            Year_on_year_growth_rate_dic[key] = round(
                ((current_amount[key] - last_years_amount[key]) / last_years_amount[key]) * 100, 2)

    # 获取2023年目标数据
    targets_data = get_df_data(targets_2023, department_all, '组织', '营业收入')
    # 计算目标差值
    target_difference = {key: targets_data[key] - cumulative_amount.get(key, 0) for key in targets_data}
    # 目标差值非正数置为0
    target_difference = {key: target_difference[key] if target_difference[key] > 0 else 0 for key in target_difference}

    # 计算累计/目标的百分比
    percentages = {key: round((cumulative_amount[key] / targets_data[key] * 100), 2) if targets_data[key] else 0 for key
                   in cumulative_amount}

    sort = {
        'ZN事业部': 1,
        'ZX事业部': 2,
        'KC事业部': 3,
        'SZ设计院': 4,
        'BR公司': 5,
        'GC公司': 6,
        'J2公司': 7,
        'SZ公司': 8,
        'XCL公司': 9
    }

    table1 = {}
    # 遍历部门，填充数据
    for key in department_all:
        table1[key] = {'季度营业收入': current_amount[key],
                     '同比': Year_on_year_growth_rate_dic[key],
                     '本年累计营业收入': cumulative_amount[key],
                     '本年目标': targets_data[key],
                     '目标完成率': percentages[key]
                     }
    department_all = sorted(department_all, key=lambda x: sort.get(x, float('inf')))

    print(department_all)

    # 将数据转换为列表
    departments_data_list = [{'组织名称': key, '季度营业收入': table1[key]['季度营业收入'], '同比': table1[key]['同比'],
          '本年累计营业收入': table1[key]['本年累计营业收入'], '本年目标': table1[key]['本年目标'],
          '目标完成率': table1[key]['目标完成率']} for key in department_all]
    # 定义数据字典
    data = {'各组织表1数据':departments_data_list}
    # 定义最大收入部门
    data['收入最高的组织'] = max(table1, key=lambda x: table1[x]['季度营业收入'])
    # 定义目标完成率大于75%的组织
    high_target_completion_departments = {department: data for department, data in table1.items() if data['目标完成率'] > 70}
    # 定义默认文字
    high_target_report_text = '无组织完成率大于75%'
    # 判断是否有多于一个部门完成率大于75%的组织，如果有，则更新文字
    if high_target_completion_departments:
        # 初始化报告文字
        high_target_report_text = ''
        # 遍历完成较好的部门
        for key in high_target_completion_departments.keys():
            # 将部门信息添加到报告文字中
            high_target_report_text = high_target_report_text + f'{key}完成较好，同比{high_target_completion_departments[key]["同比"]}%，目标完成率为{high_target_completion_departments[key]["目标完成率"]}%，'
    # 将目标完成率大于75%的文字添加到数据字典中
    data['目标完成率大于75'] = high_target_report_text

    # 获取目标完成率 < 75 的公司
    low_target_completion_departments = {department: data for department, data in table1.items() if
                                          data['目标完成率'] < 70}
    # 如果没有，设置默认文字
    # 定义一个字符串，表示“无组织完成率小于75%”
    low_target_report_text = '无组织完成率小于75%'
    # 如果有低效完成部门，则清空low_target_report_text，并遍历low_target_completion_departments中的key
    if low_target_completion_departments:
        low_target_report_text = ''
        for key in low_target_completion_departments.keys():
            # 拼接low_target_report_text，key完成未达预期，同比xx%，目标完成率为xx%，
            low_target_report_text = low_target_report_text + f'{key}完成未达预期，同比{low_target_completion_departments[key]["同比"]}%，目标完成率为{low_target_completion_departments[key]["目标完成率"]}%，'
    # 将low_target_report_text赋值给data['目标完成率小于75']
    data['目标完成率小于75'] = low_target_report_text

    # 计算营业收入合计
    data['营业收入合计'] = round(sum([table1[key]['季度营业收入'] for key in table1.keys()]),2)
    # 计算累计营业收入合计
    data['累计营业收入合计'] = round(sum([table1[key]['本年累计营业收入'] for key in table1.keys()]),2)
    # 计算目标合计
    data['目标合计'] = round(sum([table1[key]['本年目标'] for key in table1.keys()]),2)
    # 计算目标完成率合计
    data['目标完成率合计'] = round(data['累计营业收入合计'] / data['目标合计'] * 100, 2)

    return data

# 定义一个函数，用于比较 two excel files 中某列的不同值并返回结果
def make_paragraph2(excel_file_path, detail_sheet_name,project_info_sheet_name):
    # 读取项目信息表
    project_info_df = pd.read_excel(excel_file_path, sheet_name=project_info_sheet_name, header=[0],
                                    engine='openpyxl')
    # 删除指定列'区内/区外'为空的数据
    project_info_df.dropna(subset=['区内/区外'], inplace=True)

    # 读取明细帐表
    detail_df = pd.read_excel(excel_file_path, sheet_name=detail_sheet_name, skiprows=2, header=[0, 1],
                              engine='openpyxl')
    # 删除指定列'核算维度'为空的数据
    detail_df.dropna(subset=[('核算维度', 'Unnamed: 3_level_1')], inplace=True)
    # 获取所有项目信息存到列表里
    project_info_list = detail_df[('核算维度', 'Unnamed: 3_level_1')].values.tolist()
    # 用正则表达式匹配项目信息列表里的每条数据，获取项目编号，存到新的列表里
    for item in project_info_list:
        # 用正则表达式匹配项目ID，表示匹配以“项目:”开头的，由数字和字母组成的一段字符
        pattern = r"(?<=项目:)([A-Za-z0-9-]+)"
        match = re.search(pattern, item)
        if match:
            priject_id = match.group(0)
            # 将匹配到的项目ID添加到detail_df对应项目的项目编号列里
            detail_df.loc[detail_df[('核算维度', 'Unnamed: 3_level_1')] == item, ('项目信息', '项目编号')] = priject_id
            # 从项目信息表匹配项目区域信息
            area = project_info_df.loc[project_info_df['项目编号'] == priject_id, '区内/区外']
            # 处理异常，如果匹配到项目区域，就加入‘区域’列里，匹配不到就不做任何处理
            try:
                detail_df.loc[detail_df[('核算维度', 'Unnamed: 3_level_1')] == item, ('项目信息', '区域')] = \
                area.values[0]
            except IndexError:
                pass
    # 删除指定列'项目编号'为空的数据
    detail_df.dropna(subset=[('项目信息', '项目编号')], inplace=True)
    # 删除指定列'项目编号'为空的数据
    detail_df.dropna(subset=[('项目信息', '区域')], inplace=True)
    # 按照'区内/区外'列分组，计算每个分组的'本期发生额','贷方金额'之和
    inside_the_area = detail_df.loc[detail_df[('项目信息', '区域')] == '区内', ('本期发生额', '贷方金额')].sum()
    outside_the_area = detail_df.loc[detail_df[('项目信息', '区域')] == '区外', ('本期发生额', '贷方金额')].sum()
    # 如果数据为空，就不执行下面的操作直接返回
    if not inside_the_area and not outside_the_area:
        return None
    # 初始化字典
    data = {}
    # 比较两区域金额的大小
    if inside_the_area > outside_the_area:
        data['占比较大的区域'] = '区内'
        data['占比较小的区域'] = '区外'
    else:
        data['占比较大的区域'] = '区外'
        data['占比较小的区域'] = '区内'
    # 计算占比较大的区域占比
    data['占比较大的区域占比'] = round(inside_the_area / (inside_the_area + outside_the_area) * 100, 2)
    # 计算占比较小的区域占比
    data['占比较小的区域占比'] = round(outside_the_area / (inside_the_area + outside_the_area) * 100, 2)
    # 返回结果
    return data

def make_paragraph2_bak(excel_file_path, sheet_name):
    # 读取 excel 文件中的数据
    project_info_df = pd.read_excel(excel_file_path, sheet_name=sheet_name, header=[0],
                                    engine='openpyxl')
    # 删除空值
    project_info_df.dropna(subset=['区内/区外'], inplace=True)
    # 按某一列分组并计算每组的和
    grouped = project_info_df.groupby('区内/区外')['合同金额(元)'].sum()
    # 获取分组后的数据
    inside_the_area = float(grouped.get('区内', 0))
    outside_the_area = float(grouped.get('区外', 0))
    # 初始化字典
    data = {}
    # 比较两区域金额的大小
    if inside_the_area > outside_the_area:
        data['占比较大的区域'] = '区内'
        data['占比较小的区域'] = '区外'
    else:
        data['占比较大的区域'] = '区外'
        data['占比较小的区域'] = '区内'
    # 计算占比较大的区域占比
    data['占比较大的区域占比'] = round(inside_the_area / (inside_the_area + outside_the_area) * 100, 2)
    # 计算占比较小的区域占比
    data['占比较小的区域占比'] = round(outside_the_area / (inside_the_area + outside_the_area) * 100, 2)
    # 返回结果
    return data


def make_paragraph3(excel_file_path, sheet_name,total_operating_income):
    customer_info_df = pd.read_excel(excel_file_path, sheet_name=sheet_name, header=[0,1],skiprows=2,
                                    engine='openpyxl')

    # customer_info_df['最大本期发生额'] = customer_info_df[('本期发生额', '贷方金额')].max()
    customer_info_df.dropna(subset=[('本期发生额', '贷方金额')], inplace=True)
    # 去掉基本分类为空的行，因为此行是合计数据所在行
    customer_info_df.dropna(subset=[('基本分类', 'Unnamed: 0_level_1')], inplace=True)
    top_10 = customer_info_df.sort_values(by=('本期发生额', '贷方金额'), ascending=False).head(10)
    # 设置列名
    columns = ["基本分类", "企业性质分类", "商务伙伴", "科目编码", "科目名称", "年初余额_借方金额", "年初余额_贷方金额",
               "年初余额_借方金额","年初余额_贷方金额","本期发生额_借方金额", "本期发生额_贷方金额", "本年累计_借方金额",
               "本年累计_贷方金额","期末余额_借方金额", "期末余额_贷方金额"]
    customer_top_10 = {}
    # 转换为以“商务伙伴”为键的字典
    customer_dict = {}
    i = 1
    for row in top_10.values:
        row_dict = {col: row[i] for i, col in enumerate(columns)}
        row_dict['排名'] = i
        i += 1
        # 将本期发生额_贷方金额四舍五入到两位小数
        row_dict['本期发生额_贷方金额'] = round(row_dict['本期发生额_贷方金额'] / 10000, 2)
        # 计算按组织汇总表中营收占比，并四舍五入到两位小数
        row_dict['按组织汇总表中营收占比'] = round(row_dict['本期发生额_贷方金额'] / total_operating_income * 100, 2)
        # 获取商务伙伴名称
        business_partner = row_dict["商务伙伴"]
        # 把当前客户字典添加到总客户字典中
        customer_dict[business_partner] = row_dict


    # 将字典转换为列表
    customer_list = [customer_dict[key] for key in customer_dict.keys()]

    data = {}
    # 定义一个字典，用来存储数据
    data['商务伙伴排名数据'] = customer_list
    # 将客户列表存储到字典中
    data['本期发生额_贷方金额合计'] = round(sum([item['本期发生额_贷方金额'] for item in customer_list]),2)
    # 将客户列表中本期发生额_贷方金额的合计数存储到字典中，并保留两位小数
    data['按组织汇总表中营收合计占比'] = round(data['本期发生额_贷方金额合计'] / total_operating_income * 100, 2)
    # 将客户列表中本期发生额_贷方金额的合计数除以总营业额，计算占比，并保留两位小数
    data['季度营业收入最高的客户'] = customer_list[0]['商务伙伴']
    # 将客户列表中的第一个客户名称存储到字典中
    data['季度营业收入最高的客户占比'] = customer_list[0]['按组织汇总表中营收占比']
    # 将客户列表中的第一个客户占比存储到字典中

    # 统计'企业性质分类'的出现次数
    categories = [item['企业性质分类'] for item in customer_list if '企业性质分类' in item]
    category_counts = Counter(categories)

    # 找出出现次数最多的类别
    most_common_category = category_counts.most_common(1)[0][0]
    data['比例最多的客户性质性质'] = most_common_category
    return data

def set_cell_background_red(cell,color):
    # 设置单元格背景颜色为红色
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:fill'), color)
    cell._tc.get_or_add_tcPr().append(shading_elm)


def set_doc_style(doc_path):
    time.sleep(1)
    doc = Document(doc_path)
    # 获取第一个表格
    table = doc.tables[0]
    # 遍历表格中的每一行，除了标题行
    for row in table.rows[1:]:  # 如果表格第一行是标题行，从第二行开始遍历
        cell = row.cells[2]  # 第三列的单元格
        try:
            # 尝试将单元格文本转换为浮点数
            value = float(cell.text.replace('%',''))
            # 如果数值小于0，则设置背景为红色'FE939E'
            if value < 0:
                set_cell_background_red(cell,'FE939E')
        except ValueError:
            # 如果转换失败，忽略该单元格
            continue

        # 遍历表格中的每一行，除了标题行
        for row in table.rows[1:]:  # 如果表格第一行是标题行，从第二行开始遍历
            cell = row.cells[5]  # 第6列的单元格
            try:
                # 尝试将单元格文本转换为浮点数
                value = float(cell.text.replace('%', ''))
                # 如果数值大于75，则设置背景为绿色'92D050'
                if value > 75:
                    set_cell_background_red(cell, '92D050')
            except ValueError:
                # 如果转换失败，忽略该单元格
                continue
    doc.save(doc_path)


def merge_dic(dic1, dic2):
    # 遍历dic1中的每一个key
    for item in dic1.keys():
        # 将dic1中的每一个value赋值给dic2中的每一个value
        dic2[item] = dic1[item]
    # 返回dic2
    return dic2

def open_file_explorer(file_path):
    system = platform.system()
    if system == "Windows":
        subprocess.Popen(["explorer", "/select,", file_path])
    elif system == "Darwin":  # macOS
        subprocess.Popen(["open", "-R", file_path])
    elif system == "Linux":
        subprocess.Popen(["xdg-open", file_path])
    else:
        print("Unsupported operating system.")


def main(excel_file_path):
    # 创建一个字典，用于存储图片
    images_dic = dict()
    # 创建一个字典，用于存储文字
    word_dic = dict()
    # 创建一个bar1图片，并将其存储在images_dic中
    images_dic['bar1'] = make_bar1(excel_file_path=excel_file_path, current_year_sheet_name='按组织汇总（2023年）',
                                   last_year_sheet_name='按组织汇总（2022年）', targets_sheet_name='2023目标')
    # 创建一个pie1图片，并将其存储在images_dic中
    images_dic['pie1'] = make_pie1(excel_file_path, detail_sheet_name='明细账',project_info_sheet_name='项目信息1')
    # 创建一个pie4图片，并将其存储在images_dic中
    images_dic['pie4'] = make_pie4(excel_file_path=excel_file_path, sheet_name='按组织汇总（2023年）')
    # 向word_dic中添加一个键值对，key为'季度'，value为'三'
    word_dic['季度'] = '三'
    # 创建一个paragraph1，并将其存储在word_dic中
    paragraph1 = make_paragraph1(excel_file_path=excel_file_path, current_year_sheet_name='按组织汇总（2023年）',
                                   last_year_sheet_name='按组织汇总（2022年）', targets_sheet_name='2023目标')
    # 获取paragraph1中'营业收入合计'的值，并将其存储在total_operating_income中
    total_operating_income = paragraph1['营业收入合计']
    # 将paragraph1中的所有键值对添加到word_dic中
    word_dic.update(paragraph1)
    # 创建一个paragraph2，并将其存储在word_dic中
    paragraph2 = make_paragraph2(excel_file_path=excel_file_path, detail_sheet_name='明细账',project_info_sheet_name='项目信息1')
    # 将paragraph2中的所有键值对添加到word_dic中
    word_dic.update(paragraph2)
    # 创建一个paragraph3，并将其存储在word_dic中
    paragraph3 = make_paragraph3(excel_file_path=excel_file_path, sheet_name='按客户汇总',total_operating_income=total_operating_income)
    # 将paragraph3中的所有键值对添加到word_dic中
    word_dic.update(paragraph3)
    # 打印word_dic
    # 创建一个文档，并将其存储在output_docx中
    output_docx = make_doc(word_dic, images_dic)
    # 设置文档的样式
    set_doc_style(output_docx)
    # 记录信息
    logger.info(f'生成报告成功, 文件保存在{output_docx}')
    open_file_explorer(output_docx)

    # make_pie1(excel_file_path, detail_sheet_name='明细账',project_info_sheet_name='项目信息1')


def rpa_run(excel_file_path):
    main(excel_file_path)
    return '运行完成'


if __name__ == '__main__':
    excel_file_path = os.path.join(BASE_PATH, '实验数据（筛选(4).xlsx')
    main(excel_file_path)
    logger.info('运行完成')
