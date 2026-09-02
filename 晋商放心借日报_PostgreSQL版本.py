# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
# -*- coding: utf-8 -*-
"""
Created on Mon Apr 20 16:07:21 2026

@author: houlingyan

【修改版本】Oracle → PostgreSQL 适配版
PostgreSQL 连接信息：
- 地址：xxx
- 端口：xxx
- 数据库：xxx
- 用户：xxx
- 密码：xxx
"""

import os
import sys
import pandas as pd
import numpy as np

# 临时修复：由于堡垒机环境 numpy 版本较新（>=1.24），移除了过时的 np.float、np.bool、np.int，
# 但堡垒机中的 openpyxl 版本较旧（<3.1.0），还在尝试调用 numpy.float。
# 为防止 openpyxl 导入报错，在导入 openpyxl 之前手动将 Python 内置类型绑定回 numpy
try:
    if not hasattr(np, 'float_'):
        np.float_ = float
    if not hasattr(np, 'int_'):
        np.int_ = int
    if not hasattr(np, 'bool_'):
        np.bool_ = bool
except Exception:
    pass

from datetime import datetime, timedelta
from typing import Dict, Any, List, Tuple, Generator
from contextlib import contextmanager

# 【修改】导入数据库驱动 - PostgreSQL + MySQL
import pymysql
import psycopg2

# 导入 openpyxl 的相关模块，用于 Excel 的样式设置和数据写入
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

# ==========================================
# Oracle → PostgreSQL 表名映射
# ==========================================
# 根据 oracle-pgsql表名映射.xlsx 和 副本oracle-pgsql表名映射.xlsx 整理
ORACLE_TO_PGSQL_TABLE_MAP = {
    # 授信相关表
    'toutiao.lc_apply_accept': 'edw_it_odm.tt_cms_lc_apply_accept_vi',
    'toutiao.lc_appl': 'edw_it_odm.tt_cms_lc_appl_vi',
    'toutiao.lc_appl_indiv': 'edw_it_odm.tt_cms_lc_appt_indiv_vi',
    'toutiao.lc_appt_rel': 'edw_it_odm.tt_cms_lc_appt_rel_vi',
    'toutiao.lc_appl_appt': 'edw_it_odm.tt_cms_lc_appl_appt_vi',
    'toutiao.lc_apply_accept': 'edw_it_odm.tt_cms_lc_apply_accept_vi',
    'toutiao.lc_appl_expand': 'edw_it_odm.tt_cms_lc_appl_expand_vi',
    
    # 用信相关表
    'toutiao.lpb_appl': 'edw_it_odm.tt_cms_lpb_appl_vi',
    'toutiao.lpb_appt_indiv': 'edw_it_odm.tt_cms_lc_appt_indiv_vi',
    'toutiao.lpb_withdraw_accept': 'edw_it_odm.tt_cms_lpb_withdraw_accept_vi',
    
    # 借据相关表
    'toutiao.lm_loan': 'edw_it_odm.tt_gls_lm_loan_vi',
    'toutiao.lm_tt_repay_item': 'edw_it_odm.tt_gls_lm_tt_repay_item_vi',
    'toutiao.lm_tt_open': 'edw_it_odm.tt_gls_lm_tt_open_vi',
    'toutiao.lm_tt_repay_pl': 'edw_it_odm.tt_gls_lm_tt_repay_plan_vi',
    
    # 风险相关表
    'toutiao.risk_dur_appl': 'edw_it_odm.tt_plrs_risk_dur_appl_vi',
    
    # 决策相关表
    'toutiao.t_re_app_process_result': 'edw_it_odm.o_ttsp_t_re_app_process_result',
    'toutiao.t_model_log': 'edw_it_odm.o_ttsp_t_model_log',
    'toutiao.t_req_model_log': 'edw_it_odm.o_ttsp_t_req_model_log',
    
    # ODS 相关表
    'toutiao.ods_re_jinshang_inside_info_di': 'edw_it_odm.o_ttsp_ods_re_jinshang_inside_info_di',
    'toutiao.ods_re_jinshang_disbt_info_di': 'edw_it_odm.o_ttsp_ods_re_jinshang_disbt_info_di',
    'toutiao.ods_re_jinshang_proamt_info_di': 'edw_it_odm.o_ttsp_ods_re_jinshang_proamt_info_di',
    'toutiao.ods_datamodel_score_dm_zj_pb_a': 'edw_it_odm.o_ttsp_ods_datamodel_score_dm_zj_pb_a',
    'toutiao.ods_datamodel_score_dm_zj_income_credit': 'edw_it_odm.o_ttsp_ods_datamodel_score_dm_zj_income_credit',
    'toutiao.ods_datamodel_personalinfo_dataaccount_supplement': 'edw_it_odm.o_ttsp_ods_datamodel_personalinfo_dataaccount_supplement',
    'toutiao.ods_datamodel_personalinfo_dataaccount': 'edw_it_odm.o_ttsp_ods_datamodel_personalinfo_dataaccount',
    
    # 流程节点表
    'toutiao.WFI_JOIN_HIS': 'edw_it_odm.tt_cms_wfi_join_his_vi',
    'toutiao.WF_COMMENTEND': 'edw_it_odm.tt_cms_wf_commentend_vi',
}


def convert_oracle_to_pgsql(sql: str) -> str:
    """
    将 Oracle SQL 转换为 PostgreSQL 语法
    """
    result = sql
    
    # 1. 替换表名
    for oracle_table, pgsql_table in ORACLE_TO_PGSQL_TABLE_MAP.items():
        result = result.replace(oracle_table, pgsql_table)
    
    # 2. Oracle 特有语法替换
    # TO_NUMBER -> 直接使用（PostgreSQL 也支持，但通常不需要）
    # result = result.replace('TO_NUMBER(', 'CAST(')  # 根据实际情况决定是否需要
    
    # 3. TO_CHAR 日期格式化 (格式基本兼容)
    # Oracle: TO_CHAR(TO_DATE(a.appl_dt, 'YYYYMMDD'), 'YYYY-MM-DD')
    # PostgreSQL: TO_CHAR(TO_DATE(a.appl_dt, 'YYYYMMDD')::DATE, 'YYYY-MM-DD')
    # PostgreSQL 需要加 ::DATE 或 ::TIMESTAMP
    result = result.replace("TO_DATE(a.appl_dt, 'YYYYMMDD')", 
                           "TO_DATE(a.appl_dt, 'YYYYMMDD')::DATE")
    
    # 4. 替换 TO_DATE 格式（保持兼容）
    result = result.replace("TO_DATE(w.crt_dt, 'YYYYMMDD')", 
                           "TO_DATE(w.crt_dt, 'YYYYMMDD')::DATE")
    
    # 5. TRUNC 日期处理（Oracle: TRUNC(sysdate) -> PostgreSQL: CURRENT_DATE）
    result = result.replace('TRUNC(SYSDATE)', 'CURRENT_DATE')
    result = result.replace('TRUNC(SYSDATE', 'CURRENT_DATE')
    
    # 6. NVL -> COALESCE
    result = result.replace('NVL(', 'COALESCE(')
    
    # 7. SYSDATE -> CURRENT_TIMESTAMP 或 CURRENT_DATE
    result = result.replace('SYSDATE', 'CURRENT_TIMESTAMP')
    
    # 8. || 字符串连接（Oracle/PostgreSQL 都支持，保持不变）
    
    # 9. DECODE -> CASE WHEN（如果有用到）
    # 这里暂不处理，取决于实际 SQL 是否有 DECODE
    
    # 10. ROWNUM -> LIMIT 或 ROW_NUMBER()
    # 如果有 ROWNUM，需要特殊处理
    
    # 11. NVL2 -> CASE WHEN（如果有用到）
    
    # 12. SUBSTR -> SUBSTRING（PostgreSQL 两个都支持，可以不改）
    
    # 13. 移除 Oracle 的 (+) 外连接语法（如果有用到，需要改写为标准 JOIN）
    
    return result


# ==========================================
# 步骤零：用户交互与日期范围处理
# ==========================================
# 全局变量：是否开启本地 Mock 模式（测试用）
MOCK_MODE = False

class DateRangeError(Exception):
    """自定义日期范围异常"""
    pass


def get_date_range() -> Tuple[str, str]:
    """
    获取并校验日期范围（纯交互式手工输入）
    运行脚本后，会在控制台(Console)弹出提示，等待用户输入。
    """
    print("="*50)
    print(" 欢迎使用晋商环境报表自动化工具 (PostgreSQL版)")
    print("="*50)
    
    # 使用 input() 阻塞程序，等待用户在控制台键盘输入
    start_str = input("请输入【开始日期】 (格式: YYYY-MM-DD，例如 2026-04-01): ").strip()
    end_str = input("请输入【结束日期】 (格式: YYYY-MM-DD，默认同开始日期): ").strip()
    
    # 细节处理：如果用户在输入结束日期时直接敲了回车（输入为空），则默认结束日期等于开始日期（即只跑一天）
    if not end_str:
        end_str = start_str
        
    # 格式校验：尝试将字符串解析为 datetime 对象，如果格式不对（如输入了20260401而不是2026-04-01），则触发异常
    try:
        start_date = datetime.strptime(start_str, "%Y-%m-%d")
        end_date = datetime.strptime(end_str, "%Y-%m-%d")
    except ValueError:
        raise DateRangeError("❌ 日期格式错误！输入的值必须严格符合 YYYY-MM-DD 格式，请重新运行脚本。")
        
    # 逻辑校验：防止用户把时间写反了
    if end_date < start_date:
        raise DateRangeError(f"❌ 逻辑错误！结束日期 ({end_str}) 不能早于开始日期 ({start_str})。")
        
    return start_str, end_str


# ==========================================
# 步骤一：架构与连接层设计
# 目标：设计统一的数据库连接工厂，支持 MySQL 和 PostgreSQL 两种数据库，
#       利用 Python 的 contextmanager（上下文管理器）确保每次查询后自动释放连接。
# ==========================================

@contextmanager
def get_db_connection(db_type: str, config: Dict[str, Any]) -> Generator[Any, None, None]:
    """
    统一数据库连接工厂（上下文管理器）
    【修改】Oracle -> PostgreSQL
    """
    if MOCK_MODE:
        # 如果开启了本地测试模式，直接返回一个虚拟连接对象（MagicMock），阻止真实的数据库连接
        print(f"[{db_type.upper()}] 正在使用虚拟数据库连接 (MOCK_MODE=True)")
        from unittest.mock import MagicMock
        yield MagicMock()
        return
        
    conn = None
    try:
        # 分支 1：处理 MySQL 连接（保持不变）
        if db_type == 'mysql':
            conn = pymysql.connect(**config)
            
        # 【修改】分支 2：处理 PostgreSQL 连接
        elif db_type == 'postgresql':
            conn = psycopg2.connect(
                host=config['host'],
                port=config['port'],
                database=config['database'],
                user=config['user'],
                password=config['password']
            )
        
        # 分支 3：处理 Oracle 连接（已废弃，保留作为参考）
        elif db_type == 'oracle':
            raise DeprecationWarning("Oracle 连接已废弃，请使用 PostgreSQL")
        else:
            raise ValueError(f"不支持的数据库类型: {db_type}")
        
        # 将连接对象交出给外部使用
        yield conn
        
    except Exception as e:
        print(f"[错误] {db_type} 数据库连接失败，请检查网络或账号密码: {e}")
        raise
    finally:
        # 无论正常执行完毕还是报错，都要确保连接被关闭，防止连接池耗尽
        if conn:
            conn.close()


# ==========================================
# 内置 SQL 查询语句
# ==========================================

# MySQL 授信进件 SQL（仅保留决策侧笔数，金额统一由 PostgreSQL 提供）
MySQL授信进件 = """
select 
    date_format(createdate, '%Y-%m-%d') as '日期',
    count(distinct seqnum) as '沃海决策授信申请数',
    null as '沃海决策授信申请金额',
    count(distinct case when admit = '1' then seqnum else null end) as '沃海决策授信通过数',
    null as '沃海决策授信通过金额'
from
  jsxjde.ods_re_jsfxjw_inside_info_di
where
    createdate >= '{start_time}'
    and createdate <= '{end_time}'
group by
  date_format(createdate, '%Y-%m-%d')
"""

# MySQL 客户分层 SQL（笔数和利率保留，金额字段暂置空，避免引用不存在字段）
MySQL客户分层 = """
with date_statistics as (
  select 
    date_format(createdate, '%Y-%m-%d') as '日期',
    count(distinct seqnum) as '决策授信申请数',
    null as '决策授信申请金额',
    count(distinct case when admit = '1' then seqnum else null end) as '决策授信通过数',
    null as '决策授信通过金额',
        count(distinct case when admit = '1' then seqnum else null end) / count(distinct seqnum) as '授信通过率_笔数',
        null as '授信通过率_金额'
  from
    jsxjde.ods_re_jsfxjw_inside_info_di
where
    createdate >= '{start_time}'
    and createdate <= '{end_time}'
  group by
    date_format(createdate, '%Y-%m-%d')
), date_custlvl_statistics as (
  select 
    date_format(createdate, '%Y-%m-%d') as '日期',
    custlvl as '客户等级',
    count(distinct seqnum) as '决策授信申请数',
    null as '决策授信申请金额',
    count(distinct case when admit = '1' then seqnum else null end) as '授信通过笔数',
    null as '授信通过金额',
        count(distinct case when admit = '1' then seqnum else null end) / count(distinct seqnum) as '授信通过率_笔数',
        null as '授信通过率_金额'
  from
    jsxjde.ods_re_jsfxjw_inside_info_di
where
    createdate >= '{start_time}'
    and createdate <= '{end_time}'
  group by
    date_format(createdate, '%Y-%m-%d'),
    custlvl
)
select
  dcs.日期,
  dcs.客户等级,
    dcs.决策授信申请数,
    dcs.决策授信申请数 / ds.决策授信申请数 as '申请数占比',
    dcs.决策授信申请金额,
    dcs.授信通过笔数,
    dcs.授信通过笔数 / ds.决策授信通过数 as '通过数占比',
    dcs.授信通过金额,
    null as '通过金额占比',
    dcs.授信通过率_笔数,
    dcs.授信通过率_金额,
    null as '授信申请户均',
    null as '授信通过户均'
from
  date_custlvl_statistics dcs
    inner join date_statistics ds
    on dcs.日期 = ds.日期
union all
select
  日期,
  '汇总',
    决策授信申请数,
    决策授信申请数 / 决策授信申请数,
    决策授信申请金额,
    决策授信通过数,
    决策授信通过数 / 决策授信通过数,
    决策授信通过金额,
    null,
    授信通过率_笔数,
    授信通过率_金额,
    null,
    null
from
  date_statistics
order by 日期, 客户等级
"""

# MySQL 支用进件 SQL（仅保留决策侧笔数，金额统一由 PostgreSQL 提供）
MySQL支用进件 = """
select 
    date_format(createdate, '%Y-%m-%d') as 日期,
    count(1) as '决策支用申请数',
    null as '决策支用申请金额',
    count(case when admit = '1' then 1 else null end) as '决策支用通过数',
    null as '决策支用通过金额'
from (
        select 
            *,
            substring_index(seqnum, '_', 1) as base_id,
            row_number() over (partition by substring_index(seqnum, '_', 1) order by createdate desc) as rn
        from
            jsxjde.ods_re_jsfxjw_disbt_info_di
        where
            ( 
                cust_id  in (
                    select cust_id from jsxjde.t_cust_info
                )
                or 
                cust_id in (
                    select cust_id from jsxjde.t_white_infos
                )
            )
            and createdate >= '{start_time}'
            and createdate <= '{end_time}'
) t_latest
where 
    rn = 1
group by
    date_format(createdate, '%Y-%m-%d')
"""

# 客户分seg分析 SQL（MySQL，保持不变）
SQL_1_SEG = """
WITH base_metrics AS (
    SELECT
        COUNT(1) AS total_app,
        SUM(CASE WHEN CAST(sys__credit_rate AS DECIMAL(10,4)) > 0 AND CAST(sys__credit_rate AS DECIMAL(10,4)) < 12.61 THEN 1 ELSE 0 END) AS app_seg1,
        SUM(CASE WHEN CAST(sys__credit_rate AS DECIMAL(10,4)) >= 12.61 AND CAST(sys__credit_rate AS DECIMAL(10,4)) < 14.41 THEN 1 ELSE 0 END) AS app_seg2,
        SUM(CASE WHEN CAST(sys__credit_rate AS DECIMAL(10,4)) >= 14.41 AND CAST(sys__credit_rate AS DECIMAL(10,4)) < 16.21 THEN 1 ELSE 0 END) AS app_seg3,
        SUM(CASE WHEN CAST(sys__credit_rate AS DECIMAL(10,4)) >= 16.21 AND CAST(sys__credit_rate AS DECIMAL(10,4)) < 18.00 THEN 1 ELSE 0 END) AS app_seg4,
        SUM(CASE WHEN CAST(sys__credit_rate AS DECIMAL(10,4)) >= 18.00 AND CAST(sys__credit_rate AS DECIMAL(10,4)) < 19.79 THEN 1 ELSE 0 END) AS app_seg5,
        SUM(CASE WHEN CAST(sys__credit_rate AS DECIMAL(10,4)) >= 19.79 THEN 1 ELSE 0 END) AS app_seg6,
        SUM(CASE WHEN admit = '1' THEN 1 ELSE 0 END) AS total_pass,
        SUM(CASE WHEN admit = '1' AND CAST(sys__credit_rate AS DECIMAL(10,4)) > 0 AND CAST(sys__credit_rate AS DECIMAL(10,4)) < 12.61 THEN 1 ELSE 0 END) AS pass_seg1,
        SUM(CASE WHEN admit = '1' AND CAST(sys__credit_rate AS DECIMAL(10,4)) >= 12.61 AND CAST(sys__credit_rate AS DECIMAL(10,4)) < 14.41 THEN 1 ELSE 0 END) AS pass_seg2,
        SUM(CASE WHEN admit = '1' AND CAST(sys__credit_rate AS DECIMAL(10,4)) >= 14.41 AND CAST(sys__credit_rate AS DECIMAL(10,4)) < 16.21 THEN 1 ELSE 0 END) AS pass_seg3,
        SUM(CASE WHEN admit = '1' AND CAST(sys__credit_rate AS DECIMAL(10,4)) >= 16.21 AND CAST(sys__credit_rate AS DECIMAL(10,4)) < 18.00 THEN 1 ELSE 0 END) AS pass_seg4,
        SUM(CASE WHEN admit = '1' AND CAST(sys__credit_rate AS DECIMAL(10,4)) >= 18.00 AND CAST(sys__credit_rate AS DECIMAL(10,4)) < 19.79 THEN 1 ELSE 0 END) AS pass_seg5,
        SUM(CASE WHEN admit = '1' AND CAST(sys__credit_rate AS DECIMAL(10,4)) >= 19.79 THEN 1 ELSE 0 END) AS pass_seg6
    FROM ods_re_jsfxjw_inside_info_di
    WHERE createDate >= '{start_time}' AND createDate <= '{end_time}'
)
SELECT '定价' AS `客户分seg`, '(0,12.61%)' AS ` `, '[12.61%,14.41%)' AS `  `, '[14.41%,16.21%)' AS `   `, '[16.21%,18.00%)' AS `    `, '[18.00%,19.79%)' AS `     `, '[19.79%,∞)' AS `      `
UNION ALL
SELECT
    '申请' AS `客户分seg`,
    CONCAT(ROUND(app_seg1 * 100.0 / NULLIF(total_app, 0), 2), '%'),
    CONCAT(ROUND(app_seg2 * 100.0 / NULLIF(total_app, 0), 2), '%'),
    CONCAT(ROUND(app_seg3 * 100.0 / NULLIF(total_app, 0), 2), '%'),
    CONCAT(ROUND(app_seg4 * 100.0 / NULLIF(total_app, 0), 2), '%'),
    CONCAT(ROUND(app_seg5 * 100.0 / NULLIF(total_app, 0), 2), '%'),
    CONCAT(ROUND(app_seg6 * 100.0 / NULLIF(total_app, 0), 2), '%')
FROM base_metrics
UNION ALL
SELECT
    '通过' AS `客户分seg`,
    CONCAT(ROUND(pass_seg1 * 100.0 / NULLIF(total_pass, 0), 2), '%'),
    CONCAT(ROUND(pass_seg2 * 100.0 / NULLIF(total_pass, 0), 2), '%'),
    CONCAT(ROUND(pass_seg3 * 100.0 / NULLIF(total_pass, 0), 2), '%'),
    CONCAT(ROUND(pass_seg4 * 100.0 / NULLIF(total_pass, 0), 2), '%'),
    CONCAT(ROUND(pass_seg5 * 100.0 / NULLIF(total_pass, 0), 2), '%'),
    CONCAT(ROUND(pass_seg6 * 100.0 / NULLIF(total_pass, 0), 2), '%')
FROM base_metrics
"""

SQL_2_HEADCOUNT = """
WITH dim_lvl AS (
    SELECT '01' AS pd_lvl UNION ALL SELECT '02' UNION ALL SELECT '03' UNION ALL
    SELECT '04' UNION ALL SELECT '05'
),
base_data AS (
    SELECT
        pd_lvl,
        CASE
            WHEN CAST(NULLIF(sys__credit_rate, '') AS DECIMAL(10,4)) > 0 AND CAST(NULLIF(sys__credit_rate, '') AS DECIMAL(10,4)) < 12.61 THEN '(0,12.61%)'
            WHEN CAST(NULLIF(sys__credit_rate, '') AS DECIMAL(10,4)) >= 12.61 AND CAST(NULLIF(sys__credit_rate, '') AS DECIMAL(10,4)) < 14.41 THEN '[12.61%,14.41%)'
            WHEN CAST(NULLIF(sys__credit_rate, '') AS DECIMAL(10,4)) >= 14.41 AND CAST(NULLIF(sys__credit_rate, '') AS DECIMAL(10,4)) < 16.21 THEN '[14.41%,16.21%)'
            WHEN CAST(NULLIF(sys__credit_rate, '') AS DECIMAL(10,4)) >= 16.21 AND CAST(NULLIF(sys__credit_rate, '') AS DECIMAL(10,4)) < 18.00 THEN '[16.21%,18.00%)'
            WHEN CAST(NULLIF(sys__credit_rate, '') AS DECIMAL(10,4)) >= 18.00 AND CAST(NULLIF(sys__credit_rate, '') AS DECIMAL(10,4)) < 19.79 THEN '[18.00%,19.79%)'
            WHEN CAST(NULLIF(sys__credit_rate, '') AS DECIMAL(10,4)) >= 19.79 THEN '[19.79%,∞)'
        END AS seg,
        COUNT(DISTINCT cust_id) AS apply_cnt,
        COUNT(DISTINCT CASE WHEN admit = '1' THEN cust_id END) AS pass_cnt
    FROM ods_re_jsfxjw_inside_info_di
    WHERE pd_lvl IN ('01', '02', '03', '04', '05')
      AND createDate >= '{start_time}' AND createDate <= '{end_time}'
    GROUP BY pd_lvl,
        CASE
            WHEN CAST(NULLIF(sys__credit_rate, '') AS DECIMAL(10,4)) > 0 AND CAST(NULLIF(sys__credit_rate, '') AS DECIMAL(10,4)) < 12.61 THEN '(0,12.61%)'
            WHEN CAST(NULLIF(sys__credit_rate, '') AS DECIMAL(10,4)) >= 12.61 AND CAST(NULLIF(sys__credit_rate, '') AS DECIMAL(10,4)) < 14.41 THEN '[12.61%,14.41%)'
            WHEN CAST(NULLIF(sys__credit_rate, '') AS DECIMAL(10,4)) >= 14.41 AND CAST(NULLIF(sys__credit_rate, '') AS DECIMAL(10,4)) < 16.21 THEN '[14.41%,16.21%)'
            WHEN CAST(NULLIF(sys__credit_rate, '') AS DECIMAL(10,4)) >= 16.21 AND CAST(NULLIF(sys__credit_rate, '') AS DECIMAL(10,4)) < 18.00 THEN '[16.21%,18.00%)'
            WHEN CAST(NULLIF(sys__credit_rate, '') AS DECIMAL(10,4)) >= 18.00 AND CAST(NULLIF(sys__credit_rate, '') AS DECIMAL(10,4)) < 19.79 THEN '[18.00%,19.79%)'
            WHEN CAST(NULLIF(sys__credit_rate, '') AS DECIMAL(10,4)) >= 19.79 THEN '[19.79%,∞)'
        END
),
pivoted AS (
    SELECT
        d.pd_lvl,
        COALESCE(SUM(CASE WHEN b.seg = '(0,12.61%)' THEN b.apply_cnt ELSE 0 END), 0) AS apply_seg1,
        COALESCE(SUM(CASE WHEN b.seg = '[12.61%,14.41%)' THEN b.apply_cnt ELSE 0 END), 0) AS apply_seg2,
        COALESCE(SUM(CASE WHEN b.seg = '[14.41%,16.21%)' THEN b.apply_cnt ELSE 0 END), 0) AS apply_seg3,
        COALESCE(SUM(CASE WHEN b.seg = '[16.21%,18.00%)' THEN b.apply_cnt ELSE 0 END), 0) AS apply_seg4,
        COALESCE(SUM(CASE WHEN b.seg = '[18.00%,19.79%)' THEN b.apply_cnt ELSE 0 END), 0) AS apply_seg5,
        COALESCE(SUM(CASE WHEN b.seg = '[19.79%,∞)' THEN b.apply_cnt ELSE 0 END), 0) AS apply_seg6,
        COALESCE(SUM(CASE WHEN b.seg = '(0,12.61%)' THEN b.pass_cnt ELSE 0 END), 0) AS pass_seg1,
        COALESCE(SUM(CASE WHEN b.seg = '[12.61%,14.41%)' THEN b.pass_cnt ELSE 0 END), 0) AS pass_seg2,
        COALESCE(SUM(CASE WHEN b.seg = '[14.41%,16.21%)' THEN b.pass_cnt ELSE 0 END), 0) AS pass_seg3,
        COALESCE(SUM(CASE WHEN b.seg = '[16.21%,18.00%)' THEN b.pass_cnt ELSE 0 END), 0) AS pass_seg4,
        COALESCE(SUM(CASE WHEN b.seg = '[18.00%,19.79%)' THEN b.pass_cnt ELSE 0 END), 0) AS pass_seg5,
        COALESCE(SUM(CASE WHEN b.seg = '[19.79%,∞)' THEN b.pass_cnt ELSE 0 END), 0) AS pass_seg6
    FROM dim_lvl d
    LEFT JOIN base_data b ON d.pd_lvl = b.pd_lvl
    GROUP BY d.pd_lvl
),
totals AS (
    SELECT
        SUM(apply_seg1 + apply_seg2 + apply_seg3 + apply_seg4 + apply_seg5 + apply_seg6) AS total_apply_all,
        SUM(pass_seg1 + pass_seg2 + pass_seg3 + pass_seg4 + pass_seg5 + pass_seg6) AS total_pass_all
    FROM pivoted
)
SELECT
    m.pd_lvl AS `等级/定价`,
    CONCAT(ROUND(m.apply_seg1 * 100.0 / NULLIF(t.total_apply_all, 0), 2), '%') AS `申请-(0,12.61%)`,
    CONCAT(ROUND(m.apply_seg2 * 100.0 / NULLIF(t.total_apply_all, 0), 2), '%') AS `申请-[12.61%,14.41%)`,
    CONCAT(ROUND(m.apply_seg3 * 100.0 / NULLIF(t.total_apply_all, 0), 2), '%') AS `申请-[14.41%,16.21%)`,
    CONCAT(ROUND(m.apply_seg4 * 100.0 / NULLIF(t.total_apply_all, 0), 2), '%') AS `申请-[16.21%,18.00%)`,
    CONCAT(ROUND(m.apply_seg5 * 100.0 / NULLIF(t.total_apply_all, 0), 2), '%') AS `申请-[18.00%,19.79%)`,
    CONCAT(ROUND(m.apply_seg6 * 100.0 / NULLIF(t.total_apply_all, 0), 2), '%') AS `申请-[19.79%,∞)`,
    CONCAT(ROUND(m.pass_seg1 * 100.0 / NULLIF(t.total_pass_all, 0), 2), '%') AS `通过-(0,12.61%)`,
    CONCAT(ROUND(m.pass_seg2 * 100.0 / NULLIF(t.total_pass_all, 0), 2), '%') AS `通过-[12.61%,14.41%)`,
    CONCAT(ROUND(m.pass_seg3 * 100.0 / NULLIF(t.total_pass_all, 0), 2), '%') AS `通过-[14.41%,16.21%)`,
    CONCAT(ROUND(m.pass_seg4 * 100.0 / NULLIF(t.total_pass_all, 0), 2), '%') AS `通过-[16.21%,18.00%)`,
    CONCAT(ROUND(m.pass_seg5 * 100.0 / NULLIF(t.total_pass_all, 0), 2), '%') AS `通过-[18.00%,19.79%)`,
    CONCAT(ROUND(m.pass_seg6 * 100.0 / NULLIF(t.total_pass_all, 0), 2), '%') AS `通过-[19.79%,∞)`
FROM pivoted m
CROSS JOIN totals t
ORDER BY m.pd_lvl
"""

SQL_3_AMOUNT = """
SELECT
    '金额字段已迁移至PostgreSQL，当前MySQL侧暂不输出' AS `等级/定价`,
    NULL AS `通过-(0,12.61%)`,
    NULL AS `通过-[12.61%,14.41%)`,
    NULL AS `通过-[14.41%,16.21%)`,
    NULL AS `通过-[16.21%,18.00%)`,
    NULL AS `通过-[18.00%,19.79%)`,
    NULL AS `通过-[19.79%,∞)`
"""

SQL_4_WEIGHTED = """
SELECT
    NULL AS `加权平均定价-申请`,
    NULL AS `加权平均定价-通过`
"""

SQL_MANAGED_CREDIT_APPL_SEQ = """
SELECT DISTINCT c.appl_seq
FROM edw_it_odm.tt_cms_lc_appl_vi c
INNER JOIN edw_it_odm.tt_cms_lc_apply_accept_vi a ON a.appl_seq = c.appl_seq
WHERE c.loan_typ = '6137'
  AND c.wf_appr_sts = '997'
  AND a.loan_typ = '6137'
  AND a.appl_dt >= '20260604'
  {managed_credit_filter}
"""

MYSQL_DECISION_PASS_APPL_SEQ = """
SELECT DISTINCT SUBSTRING_INDEX(seqnum, '_', 1) AS appl_seq
FROM jsxjde.ods_re_jsfxjw_inside_info_di
WHERE createdate >= '2026-06-04 00:00:00'
  AND admit = '1'
"""

SQL_MANAGED_CREDIT_SERIALNUMBER = """
SELECT DISTINCT la.serialnumber
FROM edw_it_odm.tt_cms_lpb_appl_vi la
INNER JOIN (
    {SQL_MANAGED_CREDIT_APPL_SEQ}
) mc ON mc.appl_seq = la.appl_seq
WHERE la.wf_appr_sts = '997'
  AND la.loan_typ = '6137'
  AND la.serialnumber IS NOT NULL
"""

MYSQL_CREDIT_AMOUNT_DETAIL = """
SELECT
    DATE_FORMAT(createdate, '%Y-%m-%d') AS 日期,
    seqnum,
    SUBSTRING_INDEX(seqnum, '_', 1) AS appl_seq,
    custlvl AS 客户等级,
    pd_lvl AS 分层等级,
    CAST(NULLIF(sys__credit_rate, '') AS DECIMAL(10,4)) AS 定价,
    admit
FROM jsxjde.ods_re_jsfxjw_inside_info_di
WHERE createdate >= '{start_time}'
  AND createdate <= '{end_time}'
"""

MYSQL_DISBURSE_AMOUNT_DETAIL = """
SELECT
    DATE_FORMAT(createdate, '%Y-%m-%d') AS 日期,
    seqnum,
    SUBSTRING_INDEX(seqnum, '_', 1) AS dn_seq,
    CAST(NULLIF(dw__credit_rate, '') AS DECIMAL(10,4)) AS 定价,
    admit,
    ROW_NUMBER() OVER (PARTITION BY SUBSTRING_INDEX(seqnum, '_', 1) ORDER BY createdate DESC) AS rn
FROM jsxjde.ods_re_jsfxjw_disbt_info_di
WHERE (
        cust_id IN (SELECT cust_id FROM jsxjde.t_cust_info)
        OR cust_id IN (SELECT cust_id FROM jsxjde.t_white_infos)
      )
  AND createdate >= '{start_time}'
  AND createdate <= '{end_time}'
"""

PGSQL_CREDIT_AMOUNT_DETAIL = """
SELECT
    appl_seq::TEXT AS appl_seq,
    apply_amt::NUMERIC AS apply_amt
FROM edw_it_odm.tt_cms_lc_apply_accept_vi
WHERE appl_dt = '{target_date_compact}'
  AND loan_typ = '6137'
"""

PGSQL_DISBURSE_AMOUNT_DETAIL = """
SELECT
    w.dn_seq::TEXT AS dn_seq,
    w.dn_amt::NUMERIC AS dn_amt,
    MAX(CASE WHEN a.wf_appr_sts = '997' THEN 1 ELSE 0 END) AS is_passed_pg
FROM edw_it_odm.tt_cms_lpb_withdraw_accept_vi w
LEFT JOIN edw_it_odm.tt_cms_lpb_appl_vi a
  ON w.dn_seq = a.dn_seq AND a.loan_typ = '6137'
WHERE SUBSTRING(REPLACE(w.crt_dt, '-', ''), 1, 8) = '{target_date_compact}'
GROUP BY w.dn_seq, w.dn_amt
"""

# =====================================================================
# PostgreSQL 新增：授信进件 行内决策分流（当逾与人行策略拦截）
# =====================================================================
PGSQL_CREDIT_POLICY_REJECT = """
SELECT 
    COUNT(DISTINCT CASE WHEN node_name LIKE '%CommonRule%' OR node_name LIKE '%ApprovalRule6137%' THEN appl_seq END) AS "当逾拒绝数_全量",
    COUNT(DISTINCT CASE WHEN node_name LIKE '%ICR%' THEN appl_seq END) AS "人行拒绝数_全量"
FROM edw_it_odm.tt_des_r_des_sas_result_vi
WHERE loan_typ = '6137'
  AND node_type = 'policyServiceNode' 
  AND finaldecision = '50'
  AND appl_typ = '01'
  AND crt_dt >= '{start_time}'
  AND crt_dt <= '{end_time}'
"""

# 新增：授信进件 行内决策分流（工作流申请与通过）
PGSQL_CREDIT_INTERNAL_DECISION = """
WITH TargetDateApps AS (
    -- 基础进件池：限制在目标日期的进件
    SELECT appl_seq, apply_amt::NUMERIC AS apply_amt
    FROM edw_it_odm.tt_cms_lc_apply_accept_vi
    WHERE appl_dt = '{target_date_compact}'
      AND loan_typ = '6137'
),
TotalAgility AS (
    SELECT DISTINCT w.pk_value AS appl_seq
    FROM edw_it_odm.tt_cms_wfi_join_his_vi w
    JOIN edw_it_odm.tt_cms_wf_commentend_vi c
      ON w.instanceid = c.instanceid
    WHERE w.appl_type = 'AGILITY' 
      AND w.pk_col = 'appl_seq'
      AND c.nodeid = '103_a30'
),
WohaiExpand AS (
    SELECT DISTINCT appl_seq
    FROM edw_it_odm.tt_cms_lc_appl_expand_vi
    WHERE key = 'challengegrade'
),
InternalApply AS (
    SELECT t.appl_seq, ta.apply_amt
    FROM TotalAgility t
    JOIN TargetDateApps ta ON t.appl_seq = ta.appl_seq
    LEFT JOIN WohaiExpand w ON t.appl_seq = w.appl_seq
    WHERE w.appl_seq IS NULL
),
InternalPass AS (
    SELECT DISTINCT c.appl_seq
    FROM edw_it_odm.tt_cms_lc_appl_vi c
    JOIN InternalApply ia ON c.appl_seq = ia.appl_seq
    WHERE c.wf_appr_sts = '997'
)
SELECT 
    COUNT(ia.appl_seq) AS "行内决策授信申请数",
    SUM(ia.apply_amt) AS "行内决策授信申请金额",
    COUNT(ip.appl_seq) AS "行内决策授信通过数",
    SUM(CASE WHEN ip.appl_seq IS NOT NULL THEN ia.apply_amt ELSE 0 END) AS "行内决策授信通过金额"
FROM InternalApply ia
LEFT JOIN InternalPass ip ON ia.appl_seq = ip.appl_seq
"""

# 【修改】Oracle 授信进件 SQL -> PostgreSQL 版本
Oracle授信进件 = """
WITH BaseApply AS (
  -- 步骤 1：锁定前置收单表的基数（渠道侧整体口径）
  SELECT
    appl_dt,
    appl_seq,
    apply_amt::NUMERIC as apply_amt
  FROM
    edw_it_odm.tt_cms_lc_apply_accept_vi
  WHERE
    appl_dt = '{target_date_compact}'
    AND loan_typ = '6137'
),
ApprResult AS (
  -- 步骤 2：全量授信在 pgsql 侧的最终通过结果（渠道侧整体口径）
  SELECT
    appl_seq,
    MAX(CASE WHEN wf_appr_sts = '997' THEN 1 ELSE 0 END) AS is_passed
  FROM
    edw_it_odm.tt_cms_lc_appl_vi c
  WHERE
    c.loan_typ = '6137'
    AND EXISTS (
      SELECT 1 FROM BaseApply ba WHERE ba.appl_seq = c.appl_seq
    )
  GROUP BY
    c.appl_seq
)
SELECT
  TO_CHAR(a.appl_dt::DATE, 'YYYY-MM-DD') AS 日期,
  COUNT(a.appl_seq) AS 全量授信申请数,
  SUM(a.apply_amt) AS 全量授信申请金额,
  SUM(CASE WHEN b.is_passed = 1 THEN 1 ELSE 0 END) AS 最终授信通过数,
  SUM(CASE WHEN b.is_passed = 1 THEN a.apply_amt ELSE 0 END) AS 最终授信通过金额,
  SUM(CASE WHEN b.is_passed = 1 THEN 1 ELSE 0 END)::NUMERIC / COUNT(a.appl_seq) AS 全流程授信通过率_笔数,
  SUM(CASE WHEN b.is_passed = 1 THEN a.apply_amt ELSE 0 END) / SUM(a.apply_amt) AS 全流程授信通过率_金额
FROM
  BaseApply a
LEFT JOIN
  ApprResult b
ON
  a.appl_seq = b.appl_seq
GROUP BY
  TO_CHAR(a.appl_dt::DATE, 'YYYY-MM-DD')
ORDER BY
  TO_CHAR(a.appl_dt::DATE, 'YYYY-MM-DD')
"""


# 【修改】Oracle 支用进件 SQL -> PostgreSQL 版本
Oracle支用进件 = """
WITH managed_credit AS (
    -- 步骤 1：仅保留自 2026-06-04 起授信决策通过且最终授信通过的授信申请
    {SQL_MANAGED_CREDIT_APPL_SEQ}
),
BaseWithdraw AS (
    -- 步骤 2：限定为上述授信申请里，在目标日发起了支用申请的记录（通过 appl_seq 关联）
    SELECT 
        REPLACE(w.crt_dt, '-', '')::TEXT as crt_dt,
        w.dn_seq,
        w.dn_amt::NUMERIC as dn_amt,
        w.appl_seq
    FROM edw_it_odm.tt_cms_lpb_withdraw_accept_vi w
    WHERE SUBSTRING(REPLACE(w.crt_dt, '-', ''), 1, 8) = '{target_date_compact}'
      AND EXISTS (
        SELECT 1 FROM managed_credit mc WHERE mc.appl_seq = w.appl_seq
      )
),
ApprResult AS (
    -- 步骤 3：对支用审批流水表进行降维预聚合
    SELECT 
        dn_seq,
        MAX(CASE WHEN wf_appr_sts = '997' THEN 1 ELSE 0 END) AS is_passed
    FROM edw_it_odm.tt_cms_lpb_appl_vi a
    WHERE a.loan_typ = '6137'
      AND EXISTS (
        SELECT 1 FROM BaseWithdraw bw WHERE bw.dn_seq = a.dn_seq
      )
    GROUP BY dn_seq
)
-- 步骤 4：主查询，安全进行左连接与通过率聚合计算
SELECT
    TO_CHAR(bw.crt_dt::DATE, 'YYYY-MM-DD') AS 日期,
    COUNT(bw.dn_seq) AS 全量支用申请数,
    SUM(bw.dn_amt) AS 全量支用申请金额,
    SUM(CASE WHEN ar.is_passed = 1 THEN 1 ELSE 0 END) AS 最终支用通过数,
    SUM(CASE WHEN ar.is_passed = 1 THEN bw.dn_amt ELSE 0 END) AS 最终支用通过金额,
    SUM(CASE WHEN ar.is_passed = 1 THEN 1 ELSE 0 END)::NUMERIC / COUNT(bw.dn_seq) AS 全流程支用通过率_笔数,
    SUM(CASE WHEN ar.is_passed = 1 THEN bw.dn_amt ELSE 0 END) / SUM(bw.dn_amt) AS 全流程支用通过率_金额
FROM BaseWithdraw bw
LEFT JOIN ApprResult ar 
  ON bw.dn_seq = ar.dn_seq
GROUP BY TO_CHAR(bw.crt_dt::DATE, 'YYYY-MM-DD')
ORDER BY TO_CHAR(bw.crt_dt::DATE, 'YYYY-MM-DD')
"""


# 【修改】Oracle 余额统计 SQL -> PostgreSQL 版本
Oracle余额统计_嵩海 = """
WITH managed_credit AS (
    {SQL_MANAGED_CREDIT_SERIALNUMBER}
)
SELECT CAST(SUM(ll.loan_os_prcp) AS NUMERIC(20, 2)) AS 余额
FROM managed_credit mc
INNER JOIN edw_it_odm.tt_gls_lm_loan_vi ll ON mc.serialnumber = ll.loan_id
WHERE ll.loan_typ = '6137'
"""

# 新增：全量余额统计，无决策过滤和日期截断限制
Oracle余额统计_全量 = """
SELECT CAST(SUM(ll.loan_os_prcp) AS NUMERIC(20, 2)) AS 余额
FROM edw_it_odm.tt_gls_lm_loan_vi ll
WHERE ll.loan_typ = '6137'
"""

Oracle月累计放款额 = """
WITH managed_credit AS (
    {SQL_MANAGED_CREDIT_SERIALNUMBER}
),
filtered_loans AS (
    SELECT 
        ll.loan_no,
        ll.cust_id,
        ll.orig_prcp,
        ll.loan_actv_dt,
        SUBSTRING(ll.loan_actv_dt, 1, 6) AS loan_month
    FROM edw_it_odm.tt_gls_lm_loan_vi ll
    INNER JOIN managed_credit mc ON mc.serialnumber = ll.loan_id
    WHERE ll.loan_typ = '6137'
      AND SUBSTRING(ll.loan_actv_dt, 1, 6) >= '{start_month}'
      AND SUBSTRING(ll.loan_actv_dt, 1, 6) <= '{end_month}'
)
SELECT 
    loan_month AS "放款月份",
    CAST(SUM(orig_prcp) AS NUMERIC(20, 2)) AS "累计放款金额"
FROM filtered_loans
GROUP BY loan_month
ORDER BY loan_month
"""

# 新增：全量月累计放款额，无决策过滤和日期截断限制
Oracle月累计放款额_放心借全量 = """
SELECT 
    SUBSTRING(ll.loan_actv_dt, 1, 6) AS "放款月份",
    CAST(SUM(orig_prcp) AS NUMERIC(20, 2)) AS "累计放款金额"
FROM edw_it_odm.tt_gls_lm_loan_vi ll
WHERE ll.loan_typ = '6137'
  AND SUBSTRING(ll.loan_actv_dt, 1, 6) >= '{start_month}'
  AND SUBSTRING(ll.loan_actv_dt, 1, 6) <= '{end_month}'
GROUP BY SUBSTRING(ll.loan_actv_dt, 1, 6)
ORDER BY "放款月份"
"""


def build_managed_credit_filter(mysql_conn, target_date: str) -> str:
    """根据 MySQL 决策通过名单，生成 PostgreSQL 可直接拼接的 appl_seq 过滤条件。"""
    mysql_pass_df = fetch_data(MYSQL_DECISION_PASS_APPL_SEQ, mysql_conn, target_date)
    if mysql_pass_df.empty or 'appl_seq' not in mysql_pass_df.columns:
        return " AND 1 = 0"

    appl_seqs = (
        mysql_pass_df['appl_seq']
        .dropna()
        .astype(str)
        .str.strip()
    )
    appl_seqs = [seq.replace("'", "''") for seq in appl_seqs if seq]
    appl_seqs = list(dict.fromkeys(appl_seqs))

    if not appl_seqs:
        return " AND 1 = 0"

    in_list = ", ".join(f"'{seq}'" for seq in appl_seqs)
    return f" AND c.appl_seq IN ({in_list})"


def build_credit_amount_outputs(mysql_conn, postgresql_conn, target_date: str):
    """跨库补齐授信金额、客户分层金额、分层分段金额。"""
    mysql_detail = fetch_data(MYSQL_CREDIT_AMOUNT_DETAIL, mysql_conn, target_date)
    pg_detail = fetch_data(PGSQL_CREDIT_AMOUNT_DETAIL, postgresql_conn, target_date)

    if mysql_detail.empty:
        empty_tier = pd.DataFrame(columns=[
            '客户等级', '决策授信申请数', '申请数占比', '决策授信申请金额',
            '授信通过笔数', '通过数占比', '授信通过金额', '通过金额占比',
            '授信通过率_笔数', '授信通过率_金额', '授信申请户均', '授信通过户均'
        ])
        empty_seg_amount = pd.DataFrame(columns=[
            '等级/定价',
            '申请-(0,12.61%)', '申请-[12.61%,14.41%)', '申请-[14.41%,16.21%)',
            '申请-[16.21%,18.00%)', '申请-[18.00%,19.79%)', '申请-[19.79%,∞)',
            '通过-(0,12.61%)', '通过-[12.61%,14.41%)', '通过-[14.41%,16.21%)',
            '通过-[16.21%,18.00%)', '通过-[18.00%,19.79%)', '通过-[19.79%,∞)'
        ])
        return pd.DataFrame(), empty_tier, empty_seg_amount, empty_seg_amount, pd.DataFrame()

    detail = mysql_detail.copy()
    detail['appl_seq'] = detail['appl_seq'].astype(str)

    if not pg_detail.empty:
        pg_detail = pg_detail.copy()
        pg_detail['appl_seq'] = pg_detail['appl_seq'].astype(str)
        detail = detail.merge(pg_detail, on='appl_seq', how='left')
    else:
        detail['apply_amt'] = pd.NA

    detail['apply_amt'] = pd.to_numeric(detail['apply_amt'], errors='coerce')
    detail['is_pass'] = detail['admit'].astype(str) == '1'

    credit_summary = pd.DataFrame({
        '日期': [target_date],
        '沃海决策授信申请金额': [detail['apply_amt'].sum(min_count=1)],
        '沃海决策授信通过金额': [detail.loc[detail['is_pass'], 'apply_amt'].sum(min_count=1)]
    })

    tier_detail = detail.copy()
    # 剔除原来的 fillna('未分层') 逻辑，保留空值或脏数据
    tier_detail['客户等级'] = tier_detail['客户等级'].astype(str).str.strip()
    # 将 -99 合并至 E 行
    tier_detail['客户等级'] = tier_detail['客户等级'].replace({'-99': 'E', '-99.0': 'E'})
    
    # 强制补齐固定分类行，防止当日因无数据导致某行丢失（如缺少 D 行）。
    # 注意：列表中不再包含 '未分层'。所有不在列表中的脏数据/空值会被 Categorical 转为 NaN。
    fixed_tier_levels = ['A', 'B', 'C', 'D', 'E']
    tier_detail['客户等级'] = pd.Categorical(tier_detail['客户等级'], categories=fixed_tier_levels, ordered=True)
    
    # dropna=True 确保那些被 Categorical 判定为 NaN（即原本是'未分层'或空值的数据）被剔除，不再单独展示一行
    grouped = tier_detail.groupby('客户等级', dropna=True, observed=False)
    tier = grouped.apply(lambda g: pd.Series({
        '决策授信申请数': g['seqnum'].nunique(),
        '决策授信申请金额': g['apply_amt'].sum(min_count=1),
        '授信通过笔数': g.loc[g['is_pass'], 'seqnum'].nunique(),
        '授信通过金额': g.loc[g['is_pass'], 'apply_amt'].sum(min_count=1)
    })).reset_index()

    total_apply_count = tier['决策授信申请数'].sum()
    total_apply_amt = tier['决策授信申请金额'].sum(min_count=1)
    total_pass_count = tier['授信通过笔数'].sum()
    total_pass_amt = tier['授信通过金额'].sum(min_count=1)

    tier['申请数占比'] = tier['决策授信申请数'] / total_apply_count if total_apply_count else None
    tier['通过数占比'] = tier['授信通过笔数'] / total_pass_count if total_pass_count else None
    tier['通过金额占比'] = tier['授信通过金额'] / total_pass_amt if pd.notna(total_pass_amt) and total_pass_amt else None
    tier['授信通过率_笔数'] = tier['授信通过笔数'] / tier['决策授信申请数'].replace(0, pd.NA)
    tier['授信通过率_金额'] = tier['授信通过金额'] / tier['决策授信申请金额'].replace(0, pd.NA)
    tier['授信申请户均'] = tier['决策授信申请金额'] / tier['决策授信申请数'].replace(0, pd.NA)
    tier['授信通过户均'] = tier['授信通过金额'] / tier['授信通过笔数'].replace(0, pd.NA)

    summary_row = pd.DataFrame([{
        '客户等级': '汇总',
        '决策授信申请数': total_apply_count,
        '申请数占比': 1 if total_apply_count else None,
        '决策授信申请金额': total_apply_amt,
        '授信通过笔数': total_pass_count,
        '通过数占比': 1 if total_pass_count else None,
        '授信通过金额': total_pass_amt,
        '通过金额占比': 1 if pd.notna(total_pass_amt) and total_pass_amt else None,
        '授信通过率_笔数': total_pass_count / total_apply_count if total_apply_count else None,
        '授信通过率_金额': total_pass_amt / total_apply_amt if pd.notna(total_apply_amt) and total_apply_amt else None,
        '授信申请户均': total_apply_amt / total_apply_count if total_apply_count else None,
        '授信通过户均': total_pass_amt / total_pass_count if total_pass_count else None,
    }])
    tier = pd.concat([tier, summary_row], ignore_index=True)

    # 与上方 SQL 分段口径保持一致：
    # (0,12.61%)、[12.61%,14.41%)、[14.41%,16.21%)、[16.21%,18.00%)、[18.00%,19.79%)、[19.79%,∞)
    labels = ['(0,12.61%)', '[12.61%,14.41%)', '[14.41%,16.21%)', '[16.21%,18.00%)', '[18.00%,19.79%)', '[19.79%,∞)']
    seg_detail = detail.copy()
    # 剔除 fillna('未分层') 逻辑
    seg_detail['分层等级'] = seg_detail['分层等级'].astype(str).str.strip()
    # 将 -99 合并至 05 行 (基于最新需求：分seg表并入05，客户分层表并入E)
    seg_detail['分层等级'] = seg_detail['分层等级'].replace({'-99': '05', '-99.0': '05'})
    
    # 强制补齐固定分层等级，防止当日无数据导致行列丢失。同样不再包含 '未分层'。
    fixed_seg_levels = ['01', '02', '03', '04', '05']
    seg_detail['分层等级'] = pd.Categorical(seg_detail['分层等级'], categories=fixed_seg_levels, ordered=True)
    
    seg_detail['定价分段'] = pd.Series(pd.NA, index=seg_detail.index, dtype='object')
    rate_series = pd.to_numeric(seg_detail['定价'], errors='coerce')
    seg_detail.loc[(rate_series > 0) & (rate_series < 12.61), '定价分段'] = '(0,12.61%)'
    seg_detail.loc[(rate_series >= 12.61) & (rate_series < 14.41), '定价分段'] = '[12.61%,14.41%)'
    seg_detail.loc[(rate_series >= 14.41) & (rate_series < 16.21), '定价分段'] = '[14.41%,16.21%)'
    seg_detail.loc[(rate_series >= 16.21) & (rate_series < 18.00), '定价分段'] = '[16.21%,18.00%)'
    seg_detail.loc[(rate_series >= 18.00) & (rate_series < 19.79), '定价分段'] = '[18.00%,19.79%)'
    seg_detail.loc[rate_series >= 19.79, '定价分段'] = '[19.79%,∞)'
    seg_pass = seg_detail[seg_detail['is_pass'] & seg_detail['定价分段'].notna()].copy()

    def _level_sort_key(level: str):
        level_str = str(level).strip()
        try:
            return (1, int(level_str), level_str)
        except ValueError:
            return (2, float('inf'), level_str)

    # 对于 Categorical 类型的 Series，.drop_duplicates() 后依然保留所有的 categories。
    # 这里通过提取 .categories 直接获得固定好的分类层级，保证行列稳定。
    level_order = list(seg_detail['分层等级'].cat.categories)
    
    amount_matrix = pd.DataFrame({'等级/定价': level_order})
    headcount_matrix = pd.DataFrame({'等级/定价': level_order})
    total_apply_cnt_all = seg_detail['seqnum'].nunique()
    total_pass_cnt_all = seg_pass['seqnum'].nunique()
    total_apply_amt_all = seg_detail['apply_amt'].sum(min_count=1)
    total_pass_amt_all = seg_pass['apply_amt'].sum(min_count=1)
    for label in labels:
        subset = seg_detail[seg_detail['定价分段'] == label]
        if subset.empty:
            headcount_matrix[f'申请-{label}'] = [None] * len(level_order)
        else:
            apply_count_map = subset.groupby('分层等级', observed=False)['seqnum'].nunique()
            headcount_matrix[f'申请-{label}'] = [
                (apply_count_map.get(level, 0) / total_apply_cnt_all) if total_apply_cnt_all else None
                for level in level_order
            ]

    for label in labels:
        subset = seg_pass[seg_pass['定价分段'] == label]
        if subset.empty:
            headcount_matrix[f'通过-{label}'] = [None] * len(level_order)
        else:
            pass_count_map = subset.groupby('分层等级', observed=False)['seqnum'].nunique()
            headcount_matrix[f'通过-{label}'] = [
                (pass_count_map.get(level, 0) / total_pass_cnt_all) if total_pass_cnt_all else None
                for level in level_order
            ]

    for label in labels:
        subset = seg_detail[seg_detail['定价分段'] == label]
        if subset.empty:
            amount_matrix[f'申请-{label}'] = [None] * len(level_order)
        else:
            apply_map = subset.groupby('分层等级', observed=False)['apply_amt'].sum(min_count=1)
            amount_matrix[f'申请-{label}'] = [
                (apply_map.get(level, 0) / total_apply_amt_all) if pd.notna(total_apply_amt_all) and total_apply_amt_all else None
                for level in level_order
            ]

    for label in labels:
        subset = seg_pass[seg_pass['定价分段'] == label]
        if subset.empty:
            amount_matrix[f'通过-{label}'] = [None] * len(level_order)
        else:
            pass_map = subset.groupby('分层等级', observed=False)['apply_amt'].sum(min_count=1)
            amount_matrix[f'通过-{label}'] = [
                (pass_map.get(level, 0) / total_pass_amt_all) if pd.notna(total_pass_amt_all) and total_pass_amt_all else None
                for level in level_order
            ]

    return credit_summary, tier, headcount_matrix, amount_matrix, detail


def build_disburse_amount_outputs(mysql_conn, postgresql_conn, target_date: str):
    """跨库补齐支用金额和加权定价。"""
    mysql_detail = fetch_data(MYSQL_DISBURSE_AMOUNT_DETAIL, mysql_conn, target_date)
    pg_detail = fetch_data(PGSQL_DISBURSE_AMOUNT_DETAIL, postgresql_conn, target_date)

    if mysql_detail.empty:
        print(f"[支用金额诊断] {target_date} MySQL明细为空")
        return pd.DataFrame(), pd.DataFrame()

    detail = mysql_detail[mysql_detail['rn'] == 1].copy()
    detail['dn_seq'] = detail['dn_seq'].astype(str).str.strip()

    mysql_total_rows = len(mysql_detail)
    mysql_rn1_rows = len(detail)
    mysql_rn1_unique = detail['dn_seq'].nunique(dropna=True)

    if not pg_detail.empty:
        pg_detail = pg_detail.copy()
        pg_detail['dn_seq'] = pg_detail['dn_seq'].astype(str).str.strip()
        pg_total_rows = len(pg_detail)
        pg_unique_rows = pg_detail['dn_seq'].nunique(dropna=True)
        pg_dup_rows = pg_total_rows - pg_unique_rows
        detail = detail.merge(pg_detail, on='dn_seq', how='left')
    else:
        pg_total_rows = 0
        pg_unique_rows = 0
        pg_dup_rows = 0
        detail['dn_amt'] = pd.NA

    detail['dn_amt'] = pd.to_numeric(detail['dn_amt'], errors='coerce')
    detail['定价'] = pd.to_numeric(detail['定价'], errors='coerce')
    
    # 【修改点】增加 PostgreSQL 放款成功状态双重校验
    detail['is_pass_mysql'] = detail['admit'].astype(str) == '1'
    if 'is_passed_pg' in detail.columns:
        detail['is_pass'] = detail['is_pass_mysql'] & (detail['is_passed_pg'] == 1)
    else:
        detail['is_pass'] = detail['is_pass_mysql']

    matched_rows = detail['dn_amt'].notna().sum()
    unmatched_rows = detail['dn_amt'].isna().sum()
    matched_unique = detail.loc[detail['dn_amt'].notna(), 'dn_seq'].nunique(dropna=True)
    unmatched_seqs = detail.loc[detail['dn_amt'].isna(), 'dn_seq'].drop_duplicates().head(20).tolist()

    print(
        f"[支用金额诊断] {target_date} | MySQL原始行数={mysql_total_rows}, rn=1后行数={mysql_rn1_rows}, rn=1后唯一dn_seq={mysql_rn1_unique} | "
        f"PG行数={pg_total_rows}, PG唯一dn_seq={pg_unique_rows}, PG重复行数={pg_dup_rows} | "
        f"merge后命中行数={matched_rows}, 未命中行数={unmatched_rows}, 命中唯一dn_seq={matched_unique}"
    )
    if unmatched_seqs:
        print(f"[支用金额诊断] 未命中的dn_seq样例(前20个): {unmatched_seqs}")

    amount_summary = pd.DataFrame({
        '日期': [target_date],
        '决策支用申请金额': [detail['dn_amt'].sum(min_count=1)],
        # 【修正】“决策支用通过金额”仍保持仅使用 MySQL admit=1 的原始口径，不带入 997
        '决策支用通过金额': [detail.loc[detail['is_pass_mysql'], 'dn_amt'].sum(min_count=1)]
    })

    total_amt = detail['dn_amt'].sum(min_count=1)
    # 这里的 pass_amt 是为了下方加权平均定价使用，它使用了 is_pass (包含997双重校验)
    pass_amt = detail.loc[detail['is_pass'], 'dn_amt'].sum(min_count=1)
    weighted = pd.DataFrame({
        '加权平均定价-申请': [(detail['dn_amt'].mul(detail['定价']).sum(min_count=1) / total_amt) / 100 if pd.notna(total_amt) and total_amt else None],
        '加权平均定价-通过': [
            (detail.loc[detail['is_pass'], 'dn_amt'].mul(detail.loc[detail['is_pass'], '定价']).sum(min_count=1) / pass_amt) / 100
            if pd.notna(pass_amt) and pass_amt else None
        ]
    })

    return amount_summary, weighted


def fetch_data(query: str, conn, target_date: str, is_balance: bool = False, start_month: str = None, end_month: str = None, managed_credit_filter: str = '') -> pd.DataFrame:
    """
    执行SQL并抓取为DataFrame
    
    核心逻辑：
    1. 接收原始 SQL 语句
    2. 针对业务中的日期过滤条件（{target_date}, {start_time}, {end_time}）进行字符串替换
    3. 连接数据库并调用 pandas.read_sql 抓取数据
    """
    # 如果读取到的SQL为空，直接返回空的DataFrame，不阻断后续流水线
    if not query.strip():
        return pd.DataFrame()
        
    # 构建多种格式的日期和时间范围字符串，适配不同的业务过滤条件
    # 格式 1: YYYY-MM-DD (例如 2026-04-08)
    # 格式 2: YYYYMMDD   (例如 20260408)
    target_date_compact = target_date.replace('-', '')
    start_time = f"{target_date} 00:00:00"
    end_time = f"{target_date} 23:59:59"
    
    # ----------------------------------------------------
    # 特殊逻辑 1: 余额快照逻辑（仅限 PostgreSQL余额统计表）
    # 要求：只有在跑批日期是 "昨天" 时才真正执行查询，其余日期直接返回空表
    # ----------------------------------------------------
    if is_balance:
        yesterday_str = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
        if target_date != yesterday_str:
            return pd.DataFrame(columns=["SUM(LOAN_OS_PRCP)"])
    
    # ----------------------------------------------------
    # 特殊逻辑 2: 月累计放款逻辑（仅限 PostgreSQL月累计放款额表）
    # ----------------------------------------------------
    if start_month and end_month:
        # 去除用户输入的横杠，如 "2025-12" -> "202512"
        start_month_compact = start_month.replace('-', '')
        end_month_compact = end_month.replace('-', '')
        query = query.replace('{start_month}', start_month_compact)\
                     .replace('{end_month}', end_month_compact)
    
    # 使用比较稳健的字符串替换方式
    # 先处理带有空格的不规范写法，将它们统一替换为标准占位符
    query = query.replace('{ target_date }', '{target_date}')\
                 .replace('{ start_time }', '{start_time}')\
                 .replace('{ end_time }', '{end_time}')\
                 .replace('{ target_date_compact }', '{target_date_compact}')
                 
    # 然后进行真正的值替换，新增对嵌套 SQL 片段和 YYYYMMDD 无横杠格式的支持
    formatted_query = query.replace('{SQL_MANAGED_CREDIT_APPL_SEQ}', SQL_MANAGED_CREDIT_APPL_SEQ)\
                          .replace('{SQL_MANAGED_CREDIT_SERIALNUMBER}', SQL_MANAGED_CREDIT_SERIALNUMBER)\
                          .replace('{managed_credit_filter}', managed_credit_filter)\
                          .replace('{target_date}', target_date)\
                          .replace('{target_date_compact}', target_date_compact)\
                          .replace('{start_time}', start_time)\
                          .replace('{end_time}', end_time)
    formatted_query = formatted_query.replace('{SQL_MANAGED_CREDIT_APPL_SEQ}', SQL_MANAGED_CREDIT_APPL_SEQ)\
                                     .replace('{managed_credit_filter}', managed_credit_filter)
    try:
        # 如果处于测试模式，返回随机伪造的 DataFrame 以测试管道能否跑通
        if MOCK_MODE:
            import numpy as np
            # 简单匹配表名特征
            if "t_white_infos" in query and "SELECT DISTINCT cust_id" in query:
                return pd.DataFrame({"cust_id": ["1000004703", "1261125783", "999999999"]})
            elif "全量授信申请数" in query:
                return pd.DataFrame({
                    "日期": [target_date],
                    "全量授信申请数": [np.random.randint(100, 200)],
                    "全量授信申请金额": [np.random.uniform(50000, 100000)],
                    "最终授信通过数": [np.random.randint(50, 100)],
                    "最终授信通过金额": [np.random.uniform(20000, 80000)],
                    "全流程授信通过率_笔数": [np.random.uniform(0.5, 0.9)],
                    "全流程授信通过率_金额": [np.random.uniform(0.5, 0.9)]
                })
            elif ("决策授信申请数" in query or "沃海决策授信申请数" in query) and "日期" in query and "全量支用申请数" not in query and "客户等级" not in query:
                return pd.DataFrame({
                    "日期": [target_date],
                    "沃海决策授信申请数": [np.random.randint(10, 100)],
                    "沃海决策授信申请金额": [np.random.uniform(10000, 50000)],
                    "沃海决策授信通过数": [np.random.randint(5, 50)],
                    "沃海决策授信通过金额": [np.random.uniform(5000, 20000)]
                })
            elif "行内决策授信申请数" in query:
                return pd.DataFrame({
                    "行内决策授信申请数": [np.random.randint(10, 100)],
                    "行内决策授信申请金额": [np.random.uniform(10000, 50000)],
                    "行内决策授信通过数": [np.random.randint(5, 50)],
                    "行内决策授信通过金额": [np.random.uniform(5000, 20000)]
                })
            elif "人行拒绝数_全量" in query:
                return pd.DataFrame({
                    "当逾拒绝数_全量": [np.random.randint(1, 10)],
                    "人行拒绝数_全量": [np.random.randint(1, 10)]
                })
            elif "客户等级" in query and "seqnum" not in query:
                return pd.DataFrame({
                    "客户等级": ["A", "B", "C"],
                    "决策授信申请数": [1344, 1198, 1221],
                    "申请数占比": ["26.33%", "23.47%", "23.92%"],
                    "决策授信申请金额": [130472700, 98048400, 47932700],
                    "授信通过笔数": [1247, 1086, 874],
                    "通过数占比": ["35.18%", "30.63%", "24.65%"],
                    "授信通过金额": [123798100, 90528400, 34746300],
                    "通过金额占比": ["46.47%", "33.98%", "13.04%"],
                    "授信通过率_笔数": ["92.78%", "90.65%", "71.58%"],
                    "授信通过率_金额": ["94.88%", "92.33%", "72.49%"],
                    "授信申请户均": [97078, 81843, 39257],
                    "授信通过户均": [99277, 83359, 39755]
                })
            elif "MYSQL_CREDIT_AMOUNT_DETAIL" in query or ("seqnum" in query and "客户等级" in query):
                import numpy as np
                return pd.DataFrame({
                    "日期": [target_date]*10,
                    "seqnum": [f"seq_{i}" for i in range(10)],
                    "appl_seq": [f"appl_{i}" for i in range(10)],
                    "客户等级": ["A", "B", "C", "-99", "未分层"] * 2,
                    "分层等级": ["01", "02", "-99", "04", "05"] * 2,
                    "定价": [10.5, 13.0, 15.0, 17.5, 20.0] * 2,
                    "admit": ["1", "1", "0", "1", "0"] * 2
                })
            elif "PGSQL_CREDIT_AMOUNT_DETAIL" in query or ("apply_amt" in query and "appl_seq" in query):
                return pd.DataFrame({
                    "appl_seq": [f"appl_{i}" for i in range(10)],
                    "apply_amt": [10000.0, 20000.0, 15000.0, 30000.0, 25000.0] * 2
                })
            elif "加权平均定价" in query and "全量支用申请数" not in query:
                return pd.DataFrame({
                    "加权平均定价-申请": ["18.00%"],
                    "加权平均定价-通过": ["15.00%"]
                })
            elif "决策支用申请数" in query and "全量支用申请数" not in query:
                return pd.DataFrame({
                    "日期": [target_date],
                    "决策支用申请数": [40],
                    "决策支用申请金额": [25000],
                    "决策支用通过数": [35],
                    "决策支用通过金额": [20000]
                })
            elif "全量支用申请数" in query:
                return pd.DataFrame({
                    "日期": [target_date],
                    "全量支用申请数": [50],
                    "全量支用申请金额": [30000],
                    "最终支用通过数": [30],
                    "最终支用通过金额": [18000],
                    "全流程支用通过率_笔数": [0.85],
                    "全流程支用通过率_金额": [0.80]
                })
            elif "客户分seg" in query:
                return pd.DataFrame({
                    "客户分seg": ["定价", "申请", "通过"],
                    " ": ["(0,12.61%)", "5.00%", "6.00%"],
                    "  ": ["[12.61%,14.41%)", "15.00%", "18.00%"],
                    "   ": ["[14.41%,16.21%)", "20.00%", "22.00%"],
                    "    ": ["[16.21%,18.00%)", "30.00%", "28.00%"],
                    "     ": ["[18.00%,19.79%)", "25.00%", "22.00%"],
                    "      ": ["[19.79%,∞)", "5.00%", "4.00%"]
                })
            elif "等级/定价" in query:
                return pd.DataFrame({
                    "等级/定价": ["01", "02", "03", "04", "05"],
                    "申请-(0,12.61%)": ["5%", "10%", "8%", "12%", "15%"],
                    "申请-[12.61%,14.41%)": ["10%", "15%", "12%", "18%", "20%"],
                    "申请-[14.41%,16.21%)": ["20%", "18%", "25%", "22%", "25%"],
                    "申请-[16.21%,18.00%)": ["30%", "28%", "25%", "20%", "18%"],
                    "申请-[18.00%,19.79%)": ["25%", "22%", "22%", "18%", "15%"],
                    "申请-[19.79%,∞)": ["10%", "7%", "8%", "10%", "7%"],
                    "通过-(0,12.61%)": ["5%", "10%", "8%", "12%", "15%"],
                    "通过-[12.61%,14.41%)": ["10%", "15%", "12%", "18%", "20%"],
                    "通过-[14.41%,16.21%)": ["20%", "18%", "25%", "22%", "25%"],
                    "通过-[16.21%,18.00%)": ["30%", "28%", "25%", "20%", "18%"],
                    "通过-[18.00%,19.79%)": ["25%", "22%", "22%", "18%", "15%"],
                    "通过-[19.79%,∞)": ["10%", "7%", "8%", "10%", "7%"]
                })
            elif "SUM(ll.loan_os_prcp)" in query.upper() or "FILTERED_APPL" in query.upper():
                return pd.DataFrame({
                    "SUM(LOAN_OS_PRCP)": [np.random.uniform(100000, 500000)]
                })
            elif "放款月份" in query or "累计放款金额" in query:
                return pd.DataFrame({
                    "放款月份": [start_month if start_month else "202604"],
                    "累计放款金额": [np.random.uniform(500000, 1000000)]
                })
            else:
                return pd.DataFrame()
        
        # 直接使用传入的复用连接进行数据抓取
        df = pd.read_sql(formatted_query, conn)
        return df
    except Exception as e:
        print(f"[错误] 数据抽取失败，请检查SQL语法或表结构: {e}")
        return pd.DataFrame()


# ==========================================
# 步骤二：数据抽取与处理逻辑
# 目标：将分布在多个本地文本文件中的 MySQL、PostgreSQL 以及 Python 脚本代码整合
#       按统一的流程执行抓取，并将 8 个结果集按指定顺序封装为列表，供后续写入使用。
# ==========================================

def execute_data_pipeline(target_date: str, mysql_conn, postgresql_conn, start_month: str = None, end_month: str = None) -> List[Tuple[str, pd.DataFrame]]:
    """
    执行数据抓取流水线并组装 8 张表的数据块
    
    返回的结构为 List[Tuple[str, DataFrame]]，例如：[("表名", df)]
    确保了数据源能以严格的先后顺序交给 Excel 写入模块。
    """
    
    # ----------------------------------------------------
    # 表 1. 授信进件 (由 MySQL 和 PostgreSQL 数据垂直拼接)
    # ----------------------------------------------------
    sql_mysql_credit = MySQL授信进件
    df_mysql_credit = fetch_data(sql_mysql_credit, mysql_conn, target_date)

    managed_credit_filter = build_managed_credit_filter(mysql_conn, target_date)

    sql_postgresql_credit = Oracle授信进件  # 复用 Oracle 授信进件变量名（已转换为 PostgreSQL 语法）
    # 注意：Oracle授信进件 不再需要 managed_credit_filter，但为了兼容 fetch_data 签名保留传递
    df_postgresql_credit = fetch_data(sql_postgresql_credit, postgresql_conn, target_date, managed_credit_filter=managed_credit_filter)
    df_credit_amounts, df_tier_amounts, df_seg_2, df_seg_3, _ = build_credit_amount_outputs(mysql_conn, postgresql_conn, target_date)
    
    # 获取当逾拒绝和人行拒绝数据（目标日期 T）
    sql_policy_reject = PGSQL_CREDIT_POLICY_REJECT
    df_policy_reject = fetch_data(sql_policy_reject, postgresql_conn, target_date)
    
    # 获取行内决策授信进件分流数据
    sql_internal_decision = PGSQL_CREDIT_INTERNAL_DECISION
    df_internal_decision = fetch_data(sql_internal_decision, postgresql_conn, target_date)
    
    # 只要有一端有数据，就使用 pd.merge 横向合并
    if not df_mysql_credit.empty or not df_postgresql_credit.empty or not df_internal_decision.empty:
        # 容错：如果某一端为空，构造空表保证 merge 不报错
        target_dt = target_date
        if not df_postgresql_credit.empty and '日期' in df_postgresql_credit.columns and len(df_postgresql_credit['日期']) > 0:
            target_dt = df_postgresql_credit['日期'].iloc[0]
        elif not df_mysql_credit.empty and '日期' in df_mysql_credit.columns and len(df_mysql_credit['日期']) > 0:
            target_dt = df_mysql_credit['日期'].iloc[0]
                   
        if df_mysql_credit.empty or '日期' not in df_mysql_credit.columns:
            df_mysql_credit = pd.DataFrame({'日期': [target_dt], '沃海决策授信申请数': [None], '沃海决策授信申请金额': [None], '沃海决策授信通过数': [None], '沃海决策授信通过金额': [None]})
        if df_postgresql_credit.empty or '日期' not in df_postgresql_credit.columns:
            df_postgresql_credit = pd.DataFrame({'日期': [target_dt], '全量授信申请数': [None], '全量授信申请金额': [None], '最终授信通过数': [None], '最终授信通过金额': [None], '全流程授信通过率_笔数': [None], '全流程授信通过率_金额': [None]})
        if df_internal_decision.empty:
            df_internal_decision = pd.DataFrame({'日期': [target_dt], '行内决策授信申请数': [0], '行内决策授信申请金额': [0.0], '行内决策授信通过数': [0], '行内决策授信通过金额': [0.0]})
        else:
            df_internal_decision['日期'] = target_dt
            # 填补可能的空值为0，防止计算报错
            for col in ['行内决策授信申请数', '行内决策授信申请金额', '行内决策授信通过数', '行内决策授信通过金额']:
                if col in df_internal_decision.columns:
                    df_internal_decision[col] = df_internal_decision[col].fillna(0)
            
        # 以日期为基准横向合并，使用 outer join 确保哪怕单边有数据也能保留
        df_credit = pd.merge(df_mysql_credit, df_postgresql_credit, on='日期', how='outer')
        df_credit = pd.merge(df_credit, df_internal_decision, on='日期', how='outer')
        
        if not df_credit_amounts.empty:
            df_credit = pd.merge(df_credit, df_credit_amounts, on='日期', how='left', suffixes=('', '_cross'))
            for col in ['沃海决策授信申请金额', '沃海决策授信通过金额']:
                cross_col = f'{col}_cross'
                if cross_col in df_credit.columns:
                    df_credit[col] = df_credit[cross_col].combine_first(df_credit[col]) if col in df_credit.columns else df_credit[cross_col]
                    df_credit = df_credit.drop(columns=[cross_col])
        
        # 将行内分流拒绝指标拼接到最后
        if not df_policy_reject.empty:
            df_credit['当逾拒绝数_全量'] = df_policy_reject['当逾拒绝数_全量'].iloc[0] if '当逾拒绝数_全量' in df_policy_reject.columns else None
            df_credit['人行拒绝数_全量'] = df_policy_reject['人行拒绝数_全量'].iloc[0] if '人行拒绝数_全量' in df_policy_reject.columns else None
        else:
            df_credit['当逾拒绝数_全量'] = None
            df_credit['人行拒绝数_全量'] = None
            
        if MOCK_MODE:
            required_cols = [
                '全量授信申请数', '全量授信申请金额',
                '沃海决策授信申请数', '沃海决策授信申请金额',
                '沃海决策授信通过数', '沃海决策授信通过金额', '最终授信通过数', '最终授信通过金额',
                '行内决策授信申请数', '行内决策授信申请金额',
                '行内决策授信通过数', '行内决策授信通过金额',
                '全流程授信通过率_笔数', '全流程授信通过率_金额'
            ]
            for col in required_cols:
                if col not in df_credit.columns:
                    df_credit[col] = None
        
        # 【3】增加决策授信通过率计算，防止分母为0导致报错
        df_credit['沃海决策授信通过率_笔数'] = df_credit['沃海决策授信通过数'] / df_credit['沃海决策授信申请数'].replace(0, pd.NA)
        df_credit['沃海决策授信通过率_金额'] = df_credit['沃海决策授信通过金额'] / df_credit['沃海决策授信申请金额'].replace(0, pd.NA)
        df_credit['行内决策授信通过率_笔数'] = df_credit['行内决策授信通过数'] / df_credit['行内决策授信申请数'].replace(0, pd.NA)
        df_credit['行内决策授信通过率_金额'] = df_credit['行内决策授信通过金额'] / df_credit['行内决策授信申请金额'].replace(0, pd.NA)
        
        # 将 pd.NA 统一转换为 None 以便写入 Excel 时更安全
        for col in ['沃海决策授信通过率_笔数', '沃海决策授信通过率_金额', '行内决策授信通过率_笔数', '行内决策授信通过率_金额']:
            df_credit[col] = df_credit[col].where(pd.notnull(df_credit[col]), None)
            
        # 显式锁定授信进件表的列顺序
        desired_columns_credit = [
            '日期', '全量授信申请数', '全量授信申请金额',
            '行内决策授信申请数', '行内决策授信申请金额',
            '行内决策授信通过数', '行内决策授信通过金额',
            '行内决策授信通过率_笔数', '行内决策授信通过率_金额',
            '最终授信通过数', '最终授信通过金额',
            '全流程授信通过率_笔数', '全流程授信通过率_金额',
            '当逾拒绝数_全量', '人行拒绝数_全量'
        ]
        df_base = df_credit.reindex(columns=desired_columns_credit)
        
        # 构建插入在下方的“沃海”标题行与数据行
        wohai_header = {col: None for col in desired_columns_credit}
        wohai_header['行内决策授信申请数'] = '沃海决策授信申请数'
        wohai_header['行内决策授信申请金额'] = '沃海决策授信申请金额'
        wohai_header['行内决策授信通过数'] = '沃海决策授信通过数'
        wohai_header['行内决策授信通过金额'] = '沃海决策授信通过金额'
        wohai_header['行内决策授信通过率_笔数'] = '沃海决策授信通过率_笔数'
        wohai_header['行内决策授信通过率_金额'] = '沃海决策授信通过率_金额'
        
        wohai_data = {col: None for col in desired_columns_credit}
        if not df_credit.empty:
            wohai_data['行内决策授信申请数'] = df_credit['沃海决策授信申请数'].iloc[0] if '沃海决策授信申请数' in df_credit.columns else None
            wohai_data['行内决策授信申请金额'] = df_credit['沃海决策授信申请金额'].iloc[0] if '沃海决策授信申请金额' in df_credit.columns else None
            wohai_data['行内决策授信通过数'] = df_credit['沃海决策授信通过数'].iloc[0] if '沃海决策授信通过数' in df_credit.columns else None
            wohai_data['行内决策授信通过金额'] = df_credit['沃海决策授信通过金额'].iloc[0] if '沃海决策授信通过金额' in df_credit.columns else None
            wohai_data['行内决策授信通过率_笔数'] = df_credit['沃海决策授信通过率_笔数'].iloc[0] if '沃海决策授信通过率_笔数' in df_credit.columns else None
            wohai_data['行内决策授信通过率_金额'] = df_credit['沃海决策授信通过率_金额'].iloc[0] if '沃海决策授信通过率_金额' in df_credit.columns else None
            
        df_credit = pd.concat([df_base, pd.DataFrame([wohai_header, wohai_data])], ignore_index=True)
    else:
        df_credit = pd.DataFrame()
    
    # ----------------------------------------------------
    # 表 2. 客户分层 (跨库重算口径，人数与金额保持一致)
    # ----------------------------------------------------
    df_tier = df_tier_amounts.copy()
    if not df_tier.empty:
        desired_columns_tier = [
            '客户等级', '决策授信申请数', '申请数占比', '决策授信申请金额',
            '授信通过笔数', '通过数占比', '授信通过金额', '通过金额占比',
            '授信通过率_笔数', '授信通过率_金额', '授信申请户均', '授信通过户均'
        ]
        df_tier = df_tier.reindex(columns=desired_columns_tier)
    
    # ----------------------------------------------------
    # 表 3-5. 客户分seg分析
    # 表3 仍走 MySQL 原SQL；表4/5 统一走跨库明细重算逻辑，确保等级动态展示且与金额口径一致
    # ----------------------------------------------------
    sql_seg_1 = SQL_1_SEG
    df_seg_1 = fetch_data(sql_seg_1, mysql_conn, target_date)
    
    # df_seg_2 已由 build_credit_amount_outputs 动态生成
    # ----------------------------------------------------
    # 表 6. 支用进件 (由 MySQL 和 PostgreSQL 数据垂直拼接)
    # ----------------------------------------------------
    sql_mysql_disburse = MySQL支用进件
    df_mysql_disburse = fetch_data(sql_mysql_disburse, mysql_conn, target_date)

    managed_credit_filter = build_managed_credit_filter(mysql_conn, target_date)

    sql_postgresql_disburse = Oracle支用进件
    df_postgresql_disburse = fetch_data(sql_postgresql_disburse, postgresql_conn, target_date, managed_credit_filter=managed_credit_filter)
    df_disburse_amounts, df_weighted = build_disburse_amount_outputs(mysql_conn, postgresql_conn, target_date)
    
    if not df_mysql_disburse.empty or not df_postgresql_disburse.empty:
        target_dt = target_date
        if not df_postgresql_disburse.empty and '日期' in df_postgresql_disburse.columns and len(df_postgresql_disburse['日期']) > 0:
            target_dt = df_postgresql_disburse['日期'].iloc[0]
        elif not df_mysql_disburse.empty and '日期' in df_mysql_disburse.columns and len(df_mysql_disburse['日期']) > 0:
            target_dt = df_mysql_disburse['日期'].iloc[0]
                   
        if df_mysql_disburse.empty or '日期' not in df_mysql_disburse.columns:
            df_mysql_disburse = pd.DataFrame({'日期': [target_dt], '决策支用申请数': [None], '决策支用申请金额': [None], '决策支用通过数': [None], '决策支用通过金额': [None]})
        if df_postgresql_disburse.empty or '日期' not in df_postgresql_disburse.columns:
            df_postgresql_disburse = pd.DataFrame({'日期': [target_dt], '全量支用申请数': [None], '全量支用申请金额': [None], '最终支用通过数': [None], '最终支用通过金额': [None], '全流程支用通过率_笔数': [None], '全流程支用通过率_金额': [None]})
            
        df_disburse = pd.merge(df_mysql_disburse, df_postgresql_disburse, on='日期', how='outer')
        if not df_disburse_amounts.empty:
            df_disburse = pd.merge(df_disburse, df_disburse_amounts, on='日期', how='left', suffixes=('', '_cross'))
            for col in ['决策支用申请金额', '决策支用通过金额']:
                cross_col = f'{col}_cross'
                if cross_col in df_disburse.columns:
                    df_disburse[col] = df_disburse[cross_col].combine_first(df_disburse[col]) if col in df_disburse.columns else df_disburse[cross_col]
                    df_disburse = df_disburse.drop(columns=[cross_col])
        
        # 融合加权平均定价 (SQL_4_WEIGHTED)
        if not df_weighted.empty:
            avg_apply_rate = df_weighted.iloc[0]['加权平均定价-申请']
            avg_pass_rate = df_weighted.iloc[0]['加权平均定价-通过']
            
            df_disburse['加权平均定价-申请'] = None
            df_disburse['加权平均定价-通过'] = None
            
            df_disburse.at[0, '加权平均定价-申请'] = avg_apply_rate
            df_disburse.at[0, '加权平均定价-通过'] = avg_pass_rate
            
        # Mock 模式补齐缺失列
        if MOCK_MODE:
            required_cols = [
                '全量支用申请数', '决策支用申请数', '全量支用申请金额', '决策支用申请金额',
                '决策支用通过数', '决策支用通过金额', '最终支用通过数', '最终支用通过金额',
                '全流程支用通过率_笔数', '全流程支用通过率_金额'
            ]
            for col in required_cols:
                if col not in df_disburse.columns:
                    df_disburse[col] = None
        
        df_disburse['决策支用通过率_笔数'] = df_disburse['决策支用通过数'] / df_disburse['决策支用申请数'].replace(0, pd.NA)
        df_disburse['决策支用通过率_金额'] = df_disburse['决策支用通过金额'] / df_disburse['决策支用申请金额'].replace(0, pd.NA)
        
        df_disburse['决策支用通过率_笔数'] = df_disburse['决策支用通过率_笔数'].where(pd.notnull(df_disburse['决策支用通过率_笔数']), None)
        df_disburse['决策支用通过率_金额'] = df_disburse['决策支用通过率_金额'].where(pd.notnull(df_disburse['决策支用通过率_金额']), None)
        
        desired_columns = [
            '日期', 
            '全量支用申请数','全量支用申请金额', '决策支用申请数',
            '决策支用申请金额', '加权平均定价-申请',
            '决策支用通过数', '决策支用通过金额', '加权平均定价-通过',
            '最终支用通过数', '最终支用通过金额',
            '决策支用通过率_笔数', '决策支用通过率_金额',
            '全流程支用通过率_笔数', '全流程支用通过率_金额'
        ]
        df_disburse = df_disburse.reindex(columns=desired_columns)
    else:
        df_disburse = pd.DataFrame()
    
    # ----------------------------------------------------
    # 表 7. 余额统计_嵩海 (纯 PostgreSQL 数据)
    # ----------------------------------------------------
    managed_credit_filter = build_managed_credit_filter(mysql_conn, target_date)

    sql_postgresql_balance_songhai = Oracle余额统计_嵩海
    df_balance_songhai = fetch_data(sql_postgresql_balance_songhai, postgresql_conn, target_date, is_balance=True, managed_credit_filter=managed_credit_filter)
    
    # ----------------------------------------------------
    # 表 8. 累计放款_嵩海 (纯 PostgreSQL 数据)
    # ----------------------------------------------------
    sql_postgresql_cum_loan = Oracle月累计放款额
    df_cum_loan = fetch_data(sql_postgresql_cum_loan, postgresql_conn, target_date, start_month=start_month, end_month=end_month, managed_credit_filter=managed_credit_filter)
    
    # ----------------------------------------------------
    # 表 9. 余额统计_放心借全量 (纯 PostgreSQL 数据)
    # ----------------------------------------------------
    sql_postgresql_balance_full = Oracle余额统计_全量
    # 注意：全量统计不需要 managed_credit_filter
    df_balance_full = fetch_data(sql_postgresql_balance_full, postgresql_conn, target_date, is_balance=True)

    # ----------------------------------------------------
    # 表 10. 累计放款_放心借全量 (纯 PostgreSQL 数据)
    # ----------------------------------------------------
    sql_postgresql_cum_loan_full = Oracle月累计放款额_放心借全量
    df_cum_loan_full = fetch_data(sql_postgresql_cum_loan_full, postgresql_conn, target_date, start_month=start_month, end_month=end_month)

    # ==========================================
    # 严格按照业务要求组装 10 张表的输出顺序
    # ==========================================
    tables_sequence = [
        ("授信进件（行内决策包含行内当逾与人行策略，沃海决策剔除当逾与人行策略，仅取调用决策引擎数据）", df_credit),
        ("客户分层", df_tier),
        ("客户分seg", df_seg_1),
        ("客户分层&分seg_人头", df_seg_2),
        ("客户分层&分seg_金额", df_seg_3),
        ("支用进件", df_disburse),
        ("余额统计_嵩海", df_balance_songhai),
        ("累计放款_嵩海", df_cum_loan),
        ("余额统计_放心借全量", df_balance_full),
        ("累计放款_放心借全量", df_cum_loan_full)
    ]
    
    return tables_sequence


# ==========================================
# 步骤三：Excel样式渲染与追加写入
# ==========================================

def write_tables_to_sheet(writer: pd.ExcelWriter, tables: List[Tuple[str, pd.DataFrame]], sheet_name: str):
    """
    将 8 个 DataFrame 垂直写入指定的 Excel Sheet 中。
    """
    wb = writer.book
    ws = wb.create_sheet(title=sheet_name)
    
    # 样式定义：按照需求，中文标题强制要求宋体、加粗、12号字
    title_font = Font(name='宋体', bold=True, size=12)
    
    # 定义全局行指针：它记录着当前数据应该写在 Excel 的第几行
    current_row = 1  
    
    for title, df in tables:
        # [动作 1]：写入中文标题
        title_cell = ws.cell(row=current_row, column=1, value=title)
        
        # 如果 df 是空的（完全没列或者没数据），为了防止 openpyxl 后续报错，我们需要跳过后续写表头和数据的过程
        if df.empty and len(df.columns) == 0:
             current_row += 2
             continue
             
        # [动作 2]：写入表头 (DataFrame 的列名)
        header_row = current_row + 1
        for col_idx, col_name in enumerate(df.columns, start=1):
            cell = ws.cell(row=header_row, column=col_idx, value=col_name)
            
        # [动作 3]：逐行写入业务数据
        data_start_row = header_row + 1
        data_rows_count = 0
        percent_columns = {
            idx for idx, col_name in enumerate(df.columns, start=1)
            if ('率' in str(col_name))
            or ('占比' in str(col_name))
            or str(col_name).startswith('申请-')
            or str(col_name).startswith('通过-')
        }
        decimal_columns = {
            idx for idx, col_name in enumerate(df.columns, start=1)
            if str(col_name) in {'加权平均定价-申请', '加权平均定价-通过'}
        }
        
        if not df.empty:
            # 解决 openpyxl 无法写入 pd.NA 的问题
            df = df.fillna(np.nan).replace({np.nan: None})
            for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=False), start=data_start_row):
                for c_idx, value in enumerate(row, start=1):
                    cell = ws.cell(row=r_idx, column=c_idx, value=value)
                    is_numeric_value = isinstance(value, (int, float)) and not isinstance(value, bool)
                    if c_idx in percent_columns and value is not None and pd.notna(value) and is_numeric_value:
                        cell.number_format = '0.00%'
                    elif c_idx in decimal_columns and value is not None and pd.notna(value) and is_numeric_value:
                        cell.number_format = '0.00%'
                data_rows_count += 1
                    
        # [动作 4]：计算自适应列宽
        for col_idx, col_name in enumerate(df.columns, start=1):
            column_letter = get_column_letter(col_idx)
            current_width = ws.column_dimensions[column_letter].width or 0
            
            max_len = len(str(col_name).encode('gbk', errors='ignore'))
            is_percent_col = (
                ('率' in str(col_name))
                or ('占比' in str(col_name))
                or str(col_name).startswith('申请-')
                or str(col_name).startswith('通过-')
            )
            is_decimal_col = str(col_name) in {'加权平均定价-申请', '加权平均定价-通过'}
            if not df.empty:
                for item in df[col_name]:
                    is_numeric_item = isinstance(item, (int, float)) and not isinstance(item, bool)
                    if is_percent_col and item is not None and pd.notna(item) and is_numeric_item:
                        display_value = f"{item:.2%}"
                    elif is_decimal_col and item is not None and pd.notna(item) and is_numeric_item:
                        display_value = f"{item:.2%}"
                    else:
                        display_value = item
                    item_len = len(str(display_value).encode('gbk', errors='ignore'))
                    if item_len > max_len:
                        max_len = item_len
            
            optimal_width = max_len * 1.2
            if optimal_width > current_width:
                ws.column_dimensions[column_letter].width = optimal_width
        
        # [动作 5]：更新全局行指针
        current_row = data_start_row + data_rows_count + 1


# ==========================================
# 主程序执行入口与测试配置
# ==========================================

def get_month_range() -> Tuple[str, str]:
    """
    获取并校验月份范围（纯交互式手工输入）
    """
    print("\n" + "-"*50)
    print(" 【月累计放款额】参数配置")
    print("-"*50)
    start_month = input("请输入【开始月份】 (格式: YYYY-MM，例如 2025-12): ").strip()
    end_month = input("请输入【结束月份】 (格式: YYYY-MM，例如 2026-04): ").strip()
    
    if not end_month:
        end_month = start_month
        
    try:
        start_date = datetime.strptime(start_month, "%Y-%m")
        end_date = datetime.strptime(end_month, "%Y-%m")
    except ValueError:
        raise DateRangeError("❌ 月份格式错误！输入的值必须严格符合 YYYY-MM 格式。")
        
    if end_date < start_date:
        raise DateRangeError(f"❌ 逻辑错误！结束月份 ({end_month}) 不能早于开始月份 ({start_month})。")
        
    return start_month, end_month


def main():
    # 【修改】数据库连接配置 - PostgreSQL
    db_configs = {
        'mysql': {
             'host': 'xxx',
             'port': 3306,
             'user': 'xxx',
             'password': 'xxx',
             'database': 'xxx',
             'charset': 'utf8mb4'
        },
        'postgresql': {
             'host': 'xxx',
             'port': 25108,
             'database': 'xxx',
             'user': 'xxx',
             'password': 'xxx'
        }
    }
    
    # 1. 尝试获取并校验日期范围
    try:
        start_str, end_str = get_date_range()
        start_month_str, end_month_str = get_month_range()
    except DateRangeError as e:
        print(f"\n[运行中止] {e}")
        sys.exit(1)
    
    # 2. 生成包含范围内每一天的 DatetimeIndex 列表
    date_list = pd.date_range(start=start_str, end=end_str)
    
    # 3. 指定输出文件的存放路径
    output_dir = r"D:\Users\sunjichang\Desktop\wangao取数"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # 4. 组装最终的输出文件绝对路径；使用实际处理日期命名
    actual_date = end_str.replace('-', '')
    base_filename = f"晋商放心借项目报表_全量处理结果_{actual_date}.xlsx"
    output_filename = os.path.join(output_dir, base_filename)
    if os.path.exists(output_filename):
        try:
            os.remove(output_filename)
        except PermissionError:
            timestamp_suffix = datetime.now().strftime("%H%M%S")
            output_filename = os.path.join(
                output_dir,
                f"晋商放心借项目报表_全量处理结果_{actual_date}_{timestamp_suffix}.xlsx"
            )
    
    print(f"\n准备生成报表，日期范围: {start_str} 至 {end_str}")
    print(f"内置 SQL 读取完毕，准备执行查询...")
    print(f"文件将保存至: {output_filename}\n")
    
    # 5. 【修改】开启 pd.ExcelWriter 上下文管理器，使用 PostgreSQL 连接
    try:
        print("正在建立全局数据库连接...")
        with get_db_connection('mysql', db_configs['mysql']) as mysql_conn, \
             get_db_connection('postgresql', db_configs['postgresql']) as postgresql_conn:
            
            with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
                
                for current_date in date_list:
                    # 转换为标准字符串，例如 "2026-04-08"
                    target_date_str = current_date.strftime("%Y-%m-%d")
                    print(f"---> 开始处理日期: {target_date_str}")
                    
                    # 抓取与组装当天的 8 张数据表，复用全局连接
                    tables = execute_data_pipeline(target_date_str, mysql_conn, postgresql_conn, start_month=start_month_str, end_month=end_month_str)
                
                    # 将当天的数据追加写入 Writer 对应的新 Sheet
                    write_tables_to_sheet(writer, tables, sheet_name=target_date_str)
                
                # 清理 pd.ExcelWriter 底层默认生成的空 "Sheet"
                if "Sheet" in writer.book.sheetnames and len(writer.book.sheetnames) > 1:
                    # 获取 "Sheet" 对象
                    std_sheet = writer.book["Sheet"]
                    # 将第一个非 "Sheet" 的工作表设为 active
                    valid_sheets = [name for name in writer.book.sheetnames if name != "Sheet"]
                    if valid_sheets:
                        writer.book.active = writer.book[valid_sheets[0]]
                    # 彻底删除默认 Sheet
                    writer.book.remove(std_sheet)
                elif "Sheet" in writer.book.sheetnames and len(writer.book.sheetnames) == 1:
                    ws = writer.book["Sheet"]
                    ws.cell(row=1, column=1, value="无数据生成")
                
        print(f"\n[成功] 所有日期数据处理完毕，报表已一次性保存至: {output_filename}")
        
    except Exception as e:
        print(f"\n[严重错误] 执行期间发生异常: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
