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
# 但 pandas/openpyxl 等老版本库可能仍会调用，导致报错。因此在运行时动态打补丁注入
try:
    if not hasattr(np, 'float'):
        np.float = float
        np.float_ = float
        np.bool = bool
        np.int = int
        np.object = object
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

# MySQL 授信进件 SQL（保持不变）
MySQL授信进件 = """
select 
    date_format(createdate, '%Y-%m-%d') as '日期',
    count(distinct seqnum) as '决策授信申请数',
    sum(sys__credit_amount) as '决策授信申请金额',
    count(distinct case when admit = '1' then seqnum else null end) as '决策授信通过数',
    sum(case when admit = '1' then sys__credit_amount else 0 end) as '决策授信通过金额'
from
  jsxjde.ods_re_jinshang_inside_info_di
where
    createdate >= '{start_time}'
    and createdate <= '{end_time}'
group by
  date_format(createdate, '%Y-%m-%d')
"""

# MySQL 客户分层 SQL（保持不变）
MySQL客户分层 = """
with date_statistics as (
  select 
    date_format(createdate, '%Y-%m-%d') as '日期',
    count(distinct seqnum) as '决策授信申请数',
    sum(sys__credit_amount) as '决策授信申请金额',
    count(distinct case when admit = '1' then seqnum else null end) as '决策授信通过数',
    sum(case when admit = '1' then sys__credit_amount else 0 end) as '决策授信通过金额',
        count(distinct case when admit = '1' then seqnum else null end) / count(distinct seqnum) as '授信通过率_笔数',
        sum(case when admit = '1' then sys__credit_amount else 0 end) / sum(sys__credit_amount) as '授信通过率_金额'
  from
    jsxjde.ods_re_jinshang_inside_info_di
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
    sum(sys__credit_amount) as '决策授信申请金额',
    count(distinct case when admit = '1' then seqnum else null end) as '授信通过笔数',
    sum(case when admit = '1' then sys__credit_amount else 0 end) as '授信通过金额',
        count(distinct case when admit = '1' then seqnum else null end) / count(distinct seqnum) as '授信通过率_笔数',
        sum(case when admit = '1' then sys__credit_amount else 0 end) / sum(sys__credit_amount) as '授信通过率_金额'
  from
    jsxjde.ods_re_jinshang_inside_info_di
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
    dcs.授信通过金额 / ds.决策授信通过金额 as '通过金额占比',
    dcs.授信通过率_笔数,
    dcs.授信通过率_金额,
    dcs.决策授信申请金额 / dcs.决策授信申请数 as '授信申请户均',
    dcs.授信通过金额 / dcs.授信通过笔数 as '授信通过户均'
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
    决策授信通过金额 / 决策授信通过金额,
    授信通过率_笔数,
    授信通过率_金额,
    决策授信申请金额 / 决策授信申请数,
    决策授信通过金额 / 决策授信通过数
from
  date_statistics
order by 日期, 客户等级
"""

# MySQL 支用进件 SQL（修改：只取明细，金额由 PG 补充）
MySQL支用进件明细 = """
select 
    date_format(createdate, '%Y-%m-%d') as 日期,
    base_id,
    admit,
    sys__rate
from (
        select 
            *,
            substring_index(seqnum, '_', 1) as base_id,
            row_number() over (partition by substring_index(seqnum, '_', 1) order by createdate desc) as rn
        from
            jsxjde.ods_re_jinshang_disbt_info_di
        where
            ( 
                cust_id  in (
                    select cust_id from jsxjde.t_cust_info
                )
                or 
                (
                    cust_id in (
                        select cust_id from jsxjde.t_white_infos
                    )
                    and createdate >= '2026-02-26'
                )
            )
            and createdate >= '{start_time}'
            and createdate <= '{end_time}'
) t_latest
where 
    rn = 1
"""

# 客户分seg分析 SQL（MySQL，保持不变）
SQL_1_SEG = """
WITH base_metrics AS (
    SELECT
        COUNT(1) AS total_app,
        SUM(CASE WHEN sys__rate = '23.58' THEN 1 ELSE 0 END) AS app_2358,
        SUM(CASE WHEN sys__rate != '23.58' OR sys__rate IS NULL THEN 1 ELSE 0 END) AS app_not_2358,
        SUM(CASE WHEN admit = '1' THEN 1 ELSE 0 END) AS total_pass,
        SUM(CASE WHEN admit = '1' AND sys__rate = '23.58' THEN 1 ELSE 0 END) AS pass_2358,
        SUM(CASE WHEN admit = '1' AND (sys__rate != '23.58' OR sys__rate IS NULL) THEN 1 ELSE 0 END) AS pass_not_2358
    FROM ods_re_jinshang_inside_info_di
    WHERE createDate >= '{start_time}' AND createDate <= '{end_time}'
)
SELECT '定价' AS `客户分seg`, '23.58%' AS ` `, '非23.58%' AS `  `
UNION ALL
SELECT
    '申请' AS `客户分seg`,
    CONCAT(ROUND(app_2358 * 100.0 / NULLIF(total_app, 0), 2), '%'),
    CONCAT(ROUND(app_not_2358 * 100.0 / NULLIF(total_app, 0), 2), '%')
FROM base_metrics
UNION ALL
SELECT
    '通过' AS `客户分seg`,
    CONCAT(ROUND(pass_2358 * 100.0 / NULLIF(total_pass, 0), 2), '%'),
    CONCAT(ROUND(pass_not_2358 * 100.0 / NULLIF(total_pass, 0), 2), '%')
FROM base_metrics
"""

SQL_2_HEADCOUNT = """
WITH dim_lvl AS (
    SELECT 'A' AS custlvl UNION ALL SELECT 'B' UNION ALL SELECT 'C' UNION ALL
    SELECT 'D' UNION ALL SELECT 'E' UNION ALL SELECT 'F'
),
base_data AS (
    SELECT
        custlvl,
        COUNT(DISTINCT CASE WHEN sys__rate = '23.58' THEN cust_id END) AS app_2358_cnt,
        COUNT(DISTINCT CASE WHEN sys__rate != '23.58' OR sys__rate IS NULL THEN cust_id END) AS app_not_2358_cnt,
        COUNT(DISTINCT CASE WHEN admit = '1' AND sys__rate = '23.58' THEN cust_id END) AS pass_2358_cnt,
        COUNT(DISTINCT CASE WHEN admit = '1' AND (sys__rate != '23.58' OR sys__rate IS NULL) THEN cust_id END) AS pass_not_2358_cnt
    FROM ods_re_jinshang_inside_info_di
    WHERE custlvl IN ('A', 'B', 'C', 'D', 'E', 'F')
      AND createDate >= '{start_time}' AND createDate <= '{end_time}'
    GROUP BY custlvl
),
merged_data AS (
    SELECT
        d.custlvl,
        COALESCE(b.app_2358_cnt, 0) AS app_2358_cnt,
        COALESCE(b.app_not_2358_cnt, 0) AS app_not_2358_cnt,
        COALESCE(b.pass_2358_cnt, 0) AS pass_2358_cnt,
        COALESCE(b.pass_not_2358_cnt, 0) AS pass_not_2358_cnt
    FROM dim_lvl d
    LEFT JOIN base_data b ON d.custlvl = b.custlvl
),
totals AS (
    SELECT
        SUM(app_2358_cnt) AS total_app_2358,
        SUM(app_not_2358_cnt) AS total_app_not_2358,
        SUM(pass_2358_cnt) AS total_pass_2358,
        SUM(pass_not_2358_cnt) AS total_pass_not_2358
    FROM merged_data
)
SELECT
    m.custlvl AS `等级/定价`,
    CONCAT(ROUND(m.app_2358_cnt * 100.0 / NULLIF(t.total_app_2358, 0), 2), '%') AS `申请-23.58%`,
    CONCAT(ROUND(m.app_not_2358_cnt * 100.0 / NULLIF(t.total_app_not_2358, 0), 2), '%') AS `申请-非23.58%`,
    CONCAT(ROUND(m.pass_2358_cnt * 100.0 / NULLIF(t.total_pass_2358, 0), 2), '%') AS `通过-23.58%`,
    CONCAT(ROUND(m.pass_not_2358_cnt * 100.0 / NULLIF(t.total_pass_not_2358, 0), 2), '%') AS `通过-非23.58%`
FROM merged_data m
CROSS JOIN totals t
ORDER BY m.custlvl
"""

SQL_3_AMOUNT = """
WITH dim_lvl AS (
    SELECT 'A' AS custlvl UNION ALL SELECT 'B' UNION ALL SELECT 'C' UNION ALL
    SELECT 'D' UNION ALL SELECT 'E' UNION ALL SELECT 'F'
),
base_amount AS (
    SELECT
        custlvl,
        SUM(CASE WHEN sys__rate = '23.58' THEN CAST(NULLIF(sys__credit_amount, '') AS DECIMAL(18,2)) ELSE 0 END) AS app_2358_amt,
        SUM(CASE WHEN sys__rate != '23.58' OR sys__rate IS NULL THEN CAST(NULLIF(sys__credit_amount, '') AS DECIMAL(18,2)) ELSE 0 END) AS app_not_2358_amt,
        SUM(CASE WHEN admit = '1' AND sys__rate = '23.58' THEN CAST(NULLIF(sys__credit_amount, '') AS DECIMAL(18,2)) ELSE 0 END) AS pass_2358_amt,
        SUM(CASE WHEN admit = '1' AND (sys__rate != '23.58' OR sys__rate IS NULL) THEN CAST(NULLIF(sys__credit_amount, '') AS DECIMAL(18,2)) ELSE 0 END) AS pass_not_2358_amt
    FROM ods_re_jinshang_inside_info_di
    WHERE custlvl IN ('A', 'B', 'C', 'D', 'E', 'F')
      AND createDate >= '{start_time}' AND createDate <= '{end_time}'
    GROUP BY custlvl
),
merged_data AS (
    SELECT
        d.custlvl,
        COALESCE(b.app_2358_amt, 0.00) AS app_2358_amt,
        COALESCE(b.app_not_2358_amt, 0.00) AS app_not_2358_amt,
        COALESCE(b.pass_2358_amt, 0.00) AS pass_2358_amt,
        COALESCE(b.pass_not_2358_amt, 0.00) AS pass_not_2358_amt
    FROM dim_lvl d
    LEFT JOIN base_amount b ON d.custlvl = b.custlvl
),
totals AS (
    SELECT
        SUM(app_2358_amt) AS total_app_2358_amt,
        SUM(app_not_2358_amt) AS total_app_not_2358_amt,
        SUM(pass_2358_amt) AS total_pass_2358_amt,
        SUM(pass_not_2358_amt) AS total_pass_not_2358_amt
    FROM merged_data
)
SELECT
    m.custlvl AS `等级/定价`,
    CONCAT(ROUND(m.app_2358_amt * 100.0 / NULLIF(t.total_app_2358_amt, 0), 2), '%') AS `申请-23.58%`,
    CONCAT(ROUND(m.app_not_2358_amt * 100.0 / NULLIF(t.total_app_not_2358_amt, 0), 2), '%') AS `申请-非23.58%`,
    CONCAT(ROUND(m.pass_2358_amt * 100.0 / NULLIF(t.total_pass_2358_amt, 0), 2), '%') AS `通过-23.58%`,
    CONCAT(ROUND(m.pass_not_2358_amt * 100.0 / NULLIF(t.total_pass_not_2358_amt, 0), 2), '%') AS `通过-非23.58%`
FROM merged_data m
CROSS JOIN totals t
ORDER BY m.custlvl
"""

# 支用金额字典拉取 SQL（PostgreSQL）
PG支用金额字典 = """
SELECT 
    w.dn_seq::varchar AS base_id, 
    MAX(w.dn_amt::NUMERIC) AS real_loan_amount,
    MAX(CASE WHEN a.wf_appr_sts = '997' THEN 1 ELSE 0 END) AS is_passed_pg
FROM edw_it_odm.tt_cms_lpb_withdraw_accept_vi w
LEFT JOIN edw_it_odm.tt_cms_lpb_appl_vi a 
  ON w.dn_seq = a.dn_seq AND a.loan_typ = '6135'
WHERE 
  -- 增加日期过滤，防止全表扫描导致内存拉爆或结果截断
  -- 放宽到目标日期的前后几天，以防止极端跨日流转的订单匹配不到
  SUBSTRING(REPLACE(w.crt_dt, '-', ''), 1, 8) >= TO_CHAR((DATE '{target_date}' - INTERVAL '3 days'), 'YYYYMMDD')
  AND SUBSTRING(REPLACE(w.crt_dt, '-', ''), 1, 8) <= TO_CHAR((DATE '{target_date}' + INTERVAL '1 days'), 'YYYYMMDD')
GROUP BY w.dn_seq
"""
SQL_4_WEIGHTED_DETAIL = """
select 
    base_id,
    admit,
    rate
from (
    SELECT
        substring_index(seqnum, '_', 1) as base_id,
        admit,
        CAST(NULLIF(sys__rate, '') AS DECIMAL(10,4)) AS rate,
        row_number() over (partition by substring_index(seqnum, '_', 1) order by createdate desc) as rn
    FROM ods_re_jinshang_disbt_info_di
    WHERE sys__rate IS NOT NULL
      AND createDate >= '{start_time}' AND createDate <= '{end_time}'
) t_latest
where rn = 1
"""

# 【修改】Oracle 授信进件 SQL -> PostgreSQL 版本
Oracle授信进件 = """
WITH BaseApply AS (
  -- 步骤 1：锁定前置收单表的基数（确定全量申请和金额底池）
  SELECT
    appl_dt,
    appl_seq,
    apply_amt::NUMERIC as apply_amt
  FROM
    edw_it_odm.tt_cms_lc_apply_accept_vi
  WHERE
    -- 动态切片：只取当前循环的一天
    appl_dt = '{target_date_compact}' 
    AND loan_typ = '6135'
    AND (
      -- 主表基于 appl_dt 进行过滤（业务限制）
      appl_dt >= '20251203' OR 
      appl_seq IN (
      '69854392', '69854449', '69854696', '69854779', '69854823', '69855166', '69855247', '69855312',
      '69855333', '69855335', '69855318', '69855338', '69855342', '69855346', '69855373', '69855375',
      '69855384', '69855385', '69855365', '69855399', '69855401', '69855404', '69855415', '69855418',
      '69855405', '69855422', '69855410', '69855433', '69855436', '69855440', '69855441', '69855444',
      '69855445', '69855431', '69855453', '69855446', '69855474', '69855459', '69855477', '69855484',
      '69855470', '69855494', '69855486', '69855489', '69855513', '69855509', '69855512', '69855515',
      '69855538', '69855521', '69855522', '69855549', '69855552', '69855553', '69855559', '69855574',
      '69855581', '69855594', '69855615', '69855619', '69855630', '69855654', '69855655', '69855657',
      '69855658', '69855637', '69855659', '69855660', '69855662', '69855669', '69855672', '69855677',
      '69855678', '69855648', '69855681', '69855683', '69855684', '69855730', '69855712', '69855731',
      '69855732', '69855737', '69855757', '69855739', '69855761', '69855741', '69855745', '69855749',
      '69855750', '69855773', '69855779', '69855770', '69855772', '69855795', '69855781', '69855782',
      '69855783', '69855798', '69855784', '69855785', '69855786', '69855792', '69855816', '69855800',
      '69855822', '69855802', '69855803', '69855804'
      )
    )
),
ApprResult AS (
  -- 步骤 2：对审批表进行降维预聚合
  SELECT
    appl_seq,
    MAX(CASE WHEN wf_appr_sts = '997' THEN 1 ELSE 0 END) AS is_passed
  FROM
    edw_it_odm.tt_cms_lc_appl_vi c
  WHERE
    -- 关键性能优化：提前过滤，避免全表扫描
    c.loan_typ = '6135'
    AND (
      c.apply_dt >= '20251203' OR c.appl_seq IN (
          '69854392', '69854449', '69854696', '69854779', '69854823', '69855166', '69855247', '69855312',
          '69855333', '69855335', '69855318', '69855338', '69855342', '69855346', '69855373', '69855375',
          '69855384', '69855385', '69855365', '69855399', '69855401', '69855404', '69855415', '69855418',
          '69855405', '69855422', '69855410', '69855433', '69855436', '69855440', '69855441', '69855444',
          '69855445', '69855431', '69855453', '69855446', '69855474', '69855459', '69855477', '69855484',
          '69855470', '69855494', '69855486', '69855489', '69855513', '69855509', '69855512', '69855515',
          '69855538', '69855521', '69855522', '69855549', '69855552', '69855553', '69855559', '69855574',
          '69855581', '69855594', '69855615', '69855619', '69855630', '69855654', '69855655', '69855657',
          '69855658', '69855637', '69855659', '69855660', '69855662', '69855669', '69855672', '69855677',
          '69855678', '69855648', '69855681', '69855683', '69855684', '69855730', '69855712', '69855731',
          '69855732', '69855737', '69855757', '69855739', '69855761', '69855741', '69855745', '69855749',
          '69855750', '69855773', '69855779', '69855770', '69855772', '69855795', '69855781', '69855782',
          '69855783', '69855798', '69855784', '69855785', '69855786', '69855792', '69855816', '69855800',
          '69855822', '69855802', '69855803', '69855804'
        )
    )
  GROUP BY
    c.appl_seq
)
-- 步骤 3：主查询，安全进行关联与统计
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
WITH ShCredit AS (
    -- 步骤 1：解决了 cust_id 到 appl_seq 的转换需求
    SELECT
      appl_seq
    FROM edw_it_odm.tt_cms_lc_appl_vi c
    WHERE c.loan_typ = '6135' 
      AND (
        c.apply_dt >= '20251203' 
        OR c.appl_seq IN (
          '69854392', '69854449', '69854696', '69854779', '69854823', '69855166', '69855247', '69855312',
          '69855333', '69855335', '69855318', '69855338', '69855342', '69855346', '69855373', '69855375',
          '69855384', '69855385', '69855365', '69855399', '69855401', '69855404', '69855415', '69855418',
          '69855405', '69855422', '69855410', '69855433', '69855436', '69855440', '69855441', '69855444',
          '69855445', '69855431', '69855453', '69855446', '69855474', '69855459', '69855477', '69855484',
          '69855470', '69855494', '69855486', '69855489', '69855513', '69855509', '69855512', '69855515',
          '69855538', '69855521', '69855522', '69855549', '69855552', '69855553', '69855559', '69855574',
          '69855581', '69855594', '69855615', '69855619', '69855630', '69855654', '69855655', '69855657',
          '69855658', '69855637', '69855659', '69855660', '69855662', '69855669', '69855672', '69855677',
          '69855678', '69855648', '69855681', '69855683', '69855684', '69855730', '69855712', '69855731',
          '69855732', '69855737', '69855757', '69855739', '69855761', '69855741', '69855745', '69855749',
          '69855750', '69855773', '69855779', '69855770', '69855772', '69855795', '69855781', '69855782',
          '69855783', '69855798', '69855784', '69855785', '69855786', '69855792', '69855816', '69855800',
          '69855822', '69855802', '69855803', '69855804'
        )
        OR ( {dynamic_cust_id_condition} )
      )
),
BaseWithdraw AS (
    -- 步骤 2：（确定全量支用申请数和金额的底池）
    SELECT 
        REPLACE(w.crt_dt, '-', '')::TEXT as crt_dt,
        w.dn_seq,
        w.dn_amt::NUMERIC as dn_amt,
        w.appl_seq
    FROM edw_it_odm.tt_cms_lpb_withdraw_accept_vi w
    WHERE SUBSTRING(REPLACE(w.crt_dt, '-', ''), 1, 8) = '{target_date_compact}'
    -- 通过 appl_seq 限制，仅保留在受众白名单内的支用进件
    AND EXISTS (
        SELECT 1 FROM ShCredit c WHERE c.appl_seq = w.appl_seq
    )
    -- 新增限制：如果该支用属于白名单客户，则其支用表的申请时间(这里取crt_dt代指支用申请时间)必须 >= '20260226'
    -- 由于 ShCredit 中保留了所有白名单客户，我们需要在支用环节卡住时间
    -- 对于非白名单的新客/特批名单，不受此时间限制（通过上面的 appl_dt >= '20251203' 等控制）
    -- 注意：动态条件里自带了 `c.cust_id IN (...)`，但在当前查询块(BaseWithdraw)的主表是 `w`，且没有 `c` 表的 JOIN！
    -- 如果直接将包含 `c.cust_id` 的条件插入这里，会导致 SQL 解析错误 `missing FROM-clause entry for table "c"`。
    -- 因此，我们必须使用 EXISTS 子查询，回到 ShCredit 中去做判定，或者利用关联表。
    AND (
        -- 分支 1：如果是白名单客户，那么其支用时间必须 >= '20260226'
        (
            EXISTS (
                SELECT 1 FROM edw_it_odm.tt_cms_lc_appl_vi c
                WHERE c.appl_seq = w.appl_seq AND ({dynamic_cust_id_condition})
            )
            AND REPLACE(SUBSTRING(w.crt_dt, 1, 10), '-', '') >= '20260226'
        )
        OR
        -- 分支 2：如果不是白名单客户，则直接放行（不受 0226 时间限制）
        NOT EXISTS (
            SELECT 1 FROM edw_it_odm.tt_cms_lc_appl_vi c
            WHERE c.appl_seq = w.appl_seq AND ({dynamic_cust_id_condition})
        )
    )
),
ApprResult AS (
    -- 步骤 3：对支用审批流水表进行降维预聚合
    SELECT 
        dn_seq,
        MAX(CASE WHEN wf_appr_sts = '997' THEN 1 ELSE 0 END) AS is_passed
    FROM edw_it_odm.tt_cms_lpb_appl_vi a
    -- 谓词下推，仅对上面筛选出的这批支用单做扫描和聚合，避免全表扫
    WHERE a.loan_typ = '6135'
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
Oracle余额统计 = """
WITH filtered_appl AS (
    SELECT 
       la.appl_seq,
       la.serialnumber
    FROM edw_it_odm.tt_cms_lpb_appl_vi la
    WHERE la.wf_appr_sts = '997'
      AND la.loan_typ = '6135'
      AND EXISTS (
          SELECT 1 FROM edw_it_odm.tt_cms_lc_appl_vi c 
          WHERE c.appl_seq = la.appl_seq
            AND c.loan_typ = '6135'
            AND (
                 c.apply_dt >= '20251203' 
                 or c.appl_seq in (
                    '69854392', '69854449', '69854696', '69854779', '69854823', '69855166', '69855247', '69855312',
                    '69855333', '69855335', '69855318', '69855338', '69855342', '69855346', '69855373', '69855375',
                    '69855384', '69855385', '69855365', '69855399', '69855401', '69855404', '69855415', '69855418',
                    '69855405', '69855422', '69855410', '69855433', '69855436', '69855440', '69855441', '69855444',
                    '69855445', '69855431', '69855453', '69855446', '69855474', '69855459', '69855477', '69855484',
                    '69855470', '69855494', '69855486', '69855489', '69855513', '69855509', '69855512', '69855515',
                    '69855538', '69855521', '69855522', '69855549', '69855552', '69855553', '69855559', '69855574',
                    '69855581', '69855594', '69855615', '69855619', '69855630', '69855654', '69855655', '69855657',
                    '69855658', '69855637', '69855659', '69855660', '69855662', '69855669', '69855672', '69855677',
                    '69855678', '69855648', '69855681', '69855683', '69855684', '69855730', '69855712', '69855731',
                    '69855732', '69855737', '69855757', '69855739', '69855761', '69855741', '69855745', '69855749',
                    '69855750', '69855773', '69855779', '69855770', '69855772', '69855795', '69855781', '69855782',
                    '69855783', '69855798', '69855784', '69855785', '69855786', '69855792', '69855816', '69855800',
                    '69855822', '69855802', '69855803', '69855804'
                 ) 
                 or ( {dynamic_cust_id_condition} AND la.crt_dt >= '20260226' )
            )
      )
)
SELECT 
    '{target_date}' AS 统计日,
    CAST(SUM(ll.loan_os_prcp) AS NUMERIC(20, 2)) AS 余额
FROM filtered_appl fa
INNER JOIN edw_it_odm.tt_gls_lm_loan_vi ll
  ON fa.serialnumber = ll.loan_id
WHERE ll.loan_typ = '6135'
"""


# 【修改】Oracle 月累计放款额 SQL -> PostgreSQL 版本
Oracle月累计放款额 = """
-- 1. 定义人群A的源头：12.3后新客 & 12.2名单客的【授信流水号 appl_seq】
WITH Target_Appl_Seq AS (
    SELECT appl_seq
    FROM edw_it_odm.tt_cms_lc_appl_vi c
    WHERE c.loan_typ = '6135'
      AND (
        c.apply_dt >= '20251203'        
        OR c.appl_seq IN (                 
      '69854392', '69854449', '69854696', '69854779', '69854823', '69855166', '69855247', '69855312',
      '69855333', '69855335', '69855318', '69855338', '69855342', '69855346', '69855373', '69855375',
      '69855384', '69855385', '69855365', '69855399', '69855401', '69855404', '69855415', '69855418',
      '69855405', '69855422', '69855410', '69855433', '69855436', '69855440', '69855441', '69855444',
      '69855445', '69855431', '69855453', '69855446', '69855474', '69855459', '69855477', '69855484',
      '69855470', '69855494', '69855486', '69855489', '69855513', '69855509', '69855512', '69855515',
      '69855538', '69855521', '69855522', '69855549', '69855552', '69855553', '69855559', '69855574',
      '69855581', '69855594', '69855615', '69855619', '69855630', '69855654', '69855655', '69855657',
      '69855658', '69855637', '69855659', '69855660', '69855662', '69855669', '69855672', '69855677',
      '69855678', '69855648', '69855681', '69855683', '69855684', '69855730', '69855712', '69855731',
      '69855732', '69855737', '69855757', '69855739', '69855761', '69855741', '69855745', '69855749',
      '69855750', '69855773', '69855779', '69855770', '69855772', '69855795', '69855781', '69855782',
      '69855783', '69855798', '69855784', '69855785', '69855786', '69855792', '69855816', '69855800',
      '69855822', '69855802', '69855803', '69855804'
        )
      )
),
-- 通过用信表，将人群A的【appl_seq】转化为【借据号 LOAN_NO】
Target_Loan_No_A AS (
    SELECT DISTINCT loan_no
    FROM edw_it_odm.tt_cms_lpb_appl_vi
    WHERE loan_typ = '6135'
      AND appl_seq IN (SELECT appl_seq FROM Target_Appl_Seq)
      AND loan_no IS NOT NULL -- 确保用信已经生成了借据号
),
Filtered_Loans AS (
    SELECT 
        LOAN_NO,
        CUST_ID,
        ORIG_PRCP,           
        LOAN_ACTV_DT,        
        SUBSTRING(LOAN_ACTV_DT, 1, 6) AS LOAN_MONTH 
    FROM edw_it_odm.tt_gls_lm_loan_vi
    WHERE 
        LOAN_TYP = '6135'
        -- 提前过滤月份 (由外部参数注入)
        AND SUBSTRING(LOAN_ACTV_DT, 1, 6) >= '{start_month}'
        AND SUBSTRING(LOAN_ACTV_DT, 1, 6) <= '{end_month}'
        AND (
            -- 人群A的特定借据（通过关联得到的特定 LOAN_NO 放行）
            (
                LOAN_NO IN (SELECT loan_no FROM Target_Loan_No_A)
            )
            OR 
            -- 白名单客户（限制客户号，且限制 2.26 之后放款）
            (   
                (   
                    {dynamic_cust_id_condition}              
                )
                AND LOAN_ACTV_DT >= '20260226' 
            )
        )
)
-- 最终汇总：按月统计
SELECT 
    LOAN_MONTH AS "放款月份",
    SUM(ORIG_PRCP) AS "累计放款金额"
FROM Filtered_Loans
GROUP BY 
    LOAN_MONTH
ORDER BY 
    LOAN_MONTH
"""
# 【新增】星选全量余额统计 SQL (6135)
Oracle余额统计_星选全量 = """
SELECT 
    '{target_date}' AS 统计日,
    CAST(SUM(ll.loan_os_prcp) AS NUMERIC(20, 2)) AS 余额
FROM edw_it_odm.tt_gls_lm_loan_vi ll
WHERE ll.loan_typ = '6135'
"""

# 【优化口径】星选全量月累计放款额 SQL (6135) - 23-25年按年合并，26年按月展示
Oracle月累计放款额_星选全量 = """
SELECT 
    CASE 
        WHEN SUBSTRING(ll.loan_actv_dt, 1, 4) IN ('2023', '2024', '2025') THEN SUBSTRING(ll.loan_actv_dt, 1, 4) || '年'
        ELSE SUBSTRING(ll.loan_actv_dt, 1, 6)
    END AS "放款时间",
    CAST(SUM(orig_prcp) AS NUMERIC(20, 2)) AS "累计放款金额"
FROM edw_it_odm.tt_gls_lm_loan_vi ll
WHERE ll.loan_typ = '6135'
  AND SUBSTRING(ll.loan_actv_dt, 1, 6) <= '{end_month}'
GROUP BY 
    CASE 
        WHEN SUBSTRING(ll.loan_actv_dt, 1, 4) IN ('2023', '2024', '2025') THEN SUBSTRING(ll.loan_actv_dt, 1, 4) || '年'
        ELSE SUBSTRING(ll.loan_actv_dt, 1, 6)
    END
ORDER BY "放款时间"
"""

PGSQL_PRICE_MONITOR = r"""
WITH credit_base AS (
    SELECT cust_id, appl_seq, TO_DATE(SUBSTRING(apply_dt, 1, 10), 'YYYY-MM-DD') AS credit_date, apply_dt
    FROM edw_it_odm.tt_cms_lc_appl_vi
    WHERE loan_typ = '6135' 
      AND wf_appr_sts = '997'
      AND (
          REPLACE(SUBSTRING(apply_dt, 1, 10), '-', '') >= '20251203'
          OR {dynamic_cust_id_condition_no_alias}
          OR appl_seq IN (
            '69854392', '69854449', '69854696', '69854779', '69854823', '69855166', '69855247', '69855312',
            '69855333', '69855335', '69855318', '69855338', '69855342', '69855346', '69855373', '69855375',
            '69855384', '69855385', '69855365', '69855399', '69855401', '69855404', '69855415', '69855418',
            '69855405', '69855422', '69855410', '69855433', '69855436', '69855440', '69855441', '69855444',
            '69855445', '69855431', '69855453', '69855446', '69855474', '69855459', '69855477', '69855484',
            '69855470', '69855494', '69855486', '69855489', '69855513', '69855509', '69855512', '69855515',
            '69855538', '69855521', '69855522', '69855549', '69855552', '69855553', '69855559', '69855574',
            '69855581', '69855594', '69855615', '69855619', '69855630', '69855654', '69855655', '69855657',
            '69855658', '69855637', '69855659', '69855660', '69855662', '69855669', '69855672', '69855677',
            '69855678', '69855648', '69855681', '69855683', '69855684', '69855730', '69855712', '69855731',
            '69855732', '69855737', '69855757', '69855739', '69855761', '69855741', '69855745', '69855749',
            '69855750', '69855773', '69855779', '69855770', '69855772', '69855795', '69855781', '69855782',
            '69855783', '69855798', '69855784', '69855785', '69855786', '69855792', '69855816', '69855800',
            '69855822', '69855802', '69855803', '69855804'
          )
      )
      -- 剔除包含清退记录的客户
      AND NOT EXISTS (
          SELECT 1 
          FROM edw_it_odm.tt_plrs_risk_dur_appl_vi risk 
          WHERE risk.cust_id = edw_it_odm.tt_cms_lc_appl_vi.cust_id
            AND risk.apply_typ = 'Clearance'
            -- 逻辑补丁：需限定清退时间在统计日及之前，避免未来数据干扰历史报表
            AND TO_DATE(SUBSTRING(risk.crt_dt, 1, 10), 'YYYY-MM-DD') <= '{target_date}'::DATE
      )
),
expand_pricing AS (
    SELECT 
        appl_seq,
        CAST(NULLIF(REGEXP_REPLACE("value", '[^0-9\.-]', '', 'g'), '') AS NUMERIC) AS exp_rate
    FROM edw_it_odm.tt_cms_lc_appl_expand_vi
    WHERE "key" = 'annualrate' AND "value" IS NOT NULL
),
customer_dedup AS (
    SELECT 
        cb.cust_id,
        ep.exp_rate AS final_rate,
        ROW_NUMBER() OVER(PARTITION BY cb.cust_id ORDER BY cb.credit_date DESC, cb.appl_seq DESC) AS latest_rn
    FROM credit_base cb
    LEFT JOIN expand_pricing ep ON cb.appl_seq::varchar = ep.appl_seq::varchar
),
target_customers AS (
    SELECT 
        cust_id,
        CASE 
            WHEN final_rate IS NOT NULL AND (CASE WHEN final_rate <= 1 THEN final_rate * 100 ELSE final_rate END) < 23.58 THEN 1 
            ELSE 0 
        END AS is_target_pricing
    FROM customer_dedup
    WHERE latest_rn = 1 
),
base_lpb_customers AS (
    SELECT 
        lpb.cust_id,
        tc.is_target_pricing,
        ('{target_date}'::DATE - MIN(TO_DATE(SUBSTRING(lpb.crt_dt, 1, 10), 'YYYY-MM-DD'))::DATE) AS drawdown_days_diff
    FROM edw_it_odm.tt_cms_lpb_appl_vi lpb
    INNER JOIN target_customers tc ON lpb.cust_id = tc.cust_id
    WHERE lpb.loan_typ = '6135'
      AND lpb.wf_appr_sts = '997'
    GROUP BY lpb.cust_id, tc.is_target_pricing
),
customer_risk_flags AS (
    SELECT 
        bc.cust_id,
        bc.is_target_pricing,
        bc.drawdown_days_diff,
        COALESCE(MAX(
            CASE 
                WHEN risk.apply_typ IN ('AutoIncreaseRate', 'INCREASE_RATE', 'DECREASE_RATE', 'Clearance') 
                     AND ('{target_date}'::DATE - TO_DATE(SUBSTRING(risk.crt_dt, 1, 10), 'YYYY-MM-DD')::DATE) <= 90
                     AND ('{target_date}'::DATE - TO_DATE(SUBSTRING(risk.crt_dt, 1, 10), 'YYYY-MM-DD')::DATE) >= 0
                THEN 1 
                ELSE 0 
            END
        ), 0) AS has_any_increase_attempt_90_days
    FROM base_lpb_customers bc
    LEFT JOIN edw_it_odm.tt_plrs_risk_dur_appl_vi risk 
        ON bc.cust_id = risk.cust_id
        AND risk.apply_typ IN ('INCREASE_RATE', 'DECREASE_RATE', 'AutoIncreaseRate', 'Clearance')
    GROUP BY bc.cust_id, bc.is_target_pricing, bc.drawdown_days_diff
)
SELECT 
    '{target_date}' AS "日期",
    (SELECT COUNT(1) FROM credit_base) AS "基础授信底池人数",
    (SELECT COUNT(1) FROM target_customers WHERE is_target_pricing = 1) AS "过滤定价小于23.58人数",
    SUM(CASE WHEN drawdown_days_diff > 8 THEN 1 ELSE 0 END) AS "分母总人数",
    
    -- 原90天应提未提客户数 -> 90天应提未提客户数（首支超8天）
    SUM(CASE WHEN drawdown_days_diff > 8 AND is_target_pricing = 1 AND has_any_increase_attempt_90_days = 0 THEN 1 ELSE 0 END) AS "90天应提未提客户数（首支超8天）",
    
    -- 新增：90天应提未提客户数占比（首支超8天）
    CONCAT(
        ROUND(
            SUM(CASE WHEN drawdown_days_diff > 8 AND is_target_pricing = 1 AND has_any_increase_attempt_90_days = 0 THEN 1 ELSE 0 END) * 100.0 / 
            NULLIF(SUM(CASE WHEN drawdown_days_diff > 8 THEN 1 ELSE 0 END), 0), 
            2
        ), '%'
    ) AS "90天应提未提客户数占比（首支超8天）",
    
    -- 原超90天应提未提客户数 -> 90天应提未提客户数（首支超90天）
    SUM(CASE WHEN drawdown_days_diff > 90 AND is_target_pricing = 1 AND has_any_increase_attempt_90_days = 0 THEN 1 ELSE 0 END) AS "90天应提未提客户数（首支超90天）",
    
    -- 原90天应提未提客户数占比 -> 90天应提未提客户数占比（首支超90天）
    CONCAT(
        ROUND(
            SUM(CASE WHEN drawdown_days_diff > 90 AND is_target_pricing = 1 AND has_any_increase_attempt_90_days = 0 THEN 1 ELSE 0 END) * 100.0 / 
            NULLIF(SUM(CASE WHEN drawdown_days_diff > 8 THEN 1 ELSE 0 END), 0), 
            2
        ), '%'
    ) AS "90天应提未提客户数占比（首支超90天）"
FROM customer_risk_flags;
"""

def fetch_data(query: str, conn, target_date: str, is_balance: bool = False, start_month: str = None, end_month: str = None) -> pd.DataFrame:
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
            return pd.DataFrame(columns=["统计日", "余额"])
    
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
                 
    # 然后进行真正的值替换，新增对 YYYYMMDD 无横杠格式的支持
    formatted_query = query.replace('{target_date}', target_date)\
                          .replace('{target_date_compact}', target_date_compact)\
                          .replace('{start_time}', start_time)\
                          .replace('{end_time}', end_time)
    try:
        # 如果处于测试模式，返回随机伪造的 DataFrame 以测试管道能否跑通
        if MOCK_MODE:
            import random
            # 简单匹配表名特征
            if "t_white_infos" in query and "SELECT DISTINCT cust_id" in query:
                return pd.DataFrame({"cust_id": ["1000004703", "1261125783", "999999999"]})
            elif "全量授信申请数" in query:
                return pd.DataFrame({
                    "日期": [target_date],
                    "全量授信申请数": [random.randint(100, 200)],
                    "全量授信申请金额": [random.uniform(50000, 100000)],
                    "最终授信通过数": [random.randint(50, 100)],
                    "最终授信通过金额": [random.uniform(20000, 80000)],
                    "全流程授信通过率_笔数": [random.uniform(0.5, 0.9)],
                    "全流程授信通过率_金额": [random.uniform(0.5, 0.9)]
                })
            elif "决策授信申请数" in query and "日期" in query and "全量支用申请数" not in query and "客户等级" not in query:
                return pd.DataFrame({
                    "日期": [target_date],
                    "决策授信申请数": [random.randint(10, 100)],
                    "决策授信申请金额": [random.uniform(10000, 50000)],
                    "决策授信通过数": [random.randint(5, 50)],
                    "决策授信通过金额": [random.uniform(5000, 20000)]
                })
            elif "客户等级" in query:
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
                    " ": ["23.58%", "10.00%", "15.00%"],
                    "  ": ["非23.58%", "90.00%", "85.00%"]
                })
            elif "等级/定价" in query:
                return pd.DataFrame({
                    "等级/定价": ["A", "B", "C"],
                    "申请-23.58%": ["10%", "20%", "30%"],
                    "申请-非23.58%": ["90%", "80%", "70%"],
                    "通过-23.58%": ["15%", "25%", "35%"],
                    "通过-非23.58%": ["85%", "75%", "65%"]
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
            elif "90天应提未提客户数" in query:
                return pd.DataFrame({
                    "日期": [target_date],
                    "90天应提未提客户数（首支超8天）": [random.randint(10, 50)],
                    "90天应提未提客户数占比（首支超8天）": [f"{random.uniform(10, 30):.2f}%"],
                    "90天应提未提客户数（首支超90天）": [random.randint(5, 20)],
                    "90天应提未提客户数占比（首支超90天）": [f"{random.uniform(5, 15):.2f}%"]
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

def get_dynamic_cust_ids(mysql_conn, col_name="c.cust_id") -> str:
    """
    从 MySQL 的 t_white_infos 表中动态获取所有的 cust_id，
    并格式化成 PostgreSQL 支持的 IN 子句条件，每 999 个拆分为一个 OR 条件，避免超长 IN 列表错误。
    """
    query = "SELECT DISTINCT cust_id FROM jsxjde.t_white_infos WHERE cust_id IS NOT NULL"
    try:
        if MOCK_MODE:
            df = pd.DataFrame({"cust_id": ["1000004703", "1261125783", "999999999"]})
        else:
            df = pd.read_sql(query, mysql_conn)
            
        if df.empty:
            return "1=0"
        
        id_list = df['cust_id'].astype(str).tolist()
        chunks = []
        chunk_size = 999
        for i in range(0, len(id_list), chunk_size):
            chunk = id_list[i:i + chunk_size]
            formatted_chunk = ", ".join(f"'{x}'" for x in chunk)
            chunks.append(f"{col_name} IN ({formatted_chunk})")
            
        return "(" + " OR ".join(chunks) + ")"
    except Exception as e:
        print(f"[警告] 从 t_white_infos 动态获取 cust_id 失败: {e}")
        return "1=0"


def execute_data_pipeline(target_date: str, mysql_conn, postgresql_conn, start_month: str = None, end_month: str = None) -> List[Tuple[str, pd.DataFrame]]:
    """
    执行数据抓取流水线并组装 8 张表的数据块
    
    返回的结构为 List[Tuple[str, DataFrame]]，例如：[("表名", df)]
    确保了数据源能以严格的先后顺序交给 Excel 写入模块。
    """
    
    # 动态获取白名单客户 ID 并生成 PostgreSQL SQL 过滤条件
    print("正在从 MySQL 动态获取 t_white_infos 白名单 cust_id...")
    dynamic_c_cust_id_cond = get_dynamic_cust_ids(mysql_conn, col_name="c.cust_id")
    dynamic_cust_id_cond = get_dynamic_cust_ids(mysql_conn, col_name="cust_id")
    
    # ----------------------------------------------------
    # 表 1. 授信进件 (由 MySQL 和 PostgreSQL 数据垂直拼接)
    # ----------------------------------------------------
    sql_mysql_credit = MySQL授信进件
    df_mysql_credit = fetch_data(sql_mysql_credit, mysql_conn, target_date)
    
    sql_postgresql_credit = Oracle授信进件  # 复用 Oracle 授信进件变量名（已转换为 PostgreSQL 语法）
    df_postgresql_credit = fetch_data(sql_postgresql_credit, postgresql_conn, target_date)
    
    # 只要有一端有数据，就使用 pd.merge 横向合并
    if not df_mysql_credit.empty or not df_postgresql_credit.empty:
        # 容错：如果某一端为空，构造空表保证 merge 不报错
        target_dt = target_date
        if not df_postgresql_credit.empty and '日期' in df_postgresql_credit.columns and len(df_postgresql_credit['日期']) > 0:
            target_dt = df_postgresql_credit['日期'].iloc[0]
        elif not df_mysql_credit.empty and '日期' in df_mysql_credit.columns and len(df_mysql_credit['日期']) > 0:
            target_dt = df_mysql_credit['日期'].iloc[0]
                   
        if df_mysql_credit.empty or '日期' not in df_mysql_credit.columns:
            df_mysql_credit = pd.DataFrame({'日期': [target_dt], '决策授信申请数': [None], '决策授信申请金额': [None], '决策授信通过数': [None], '决策授信通过金额': [None]})
        if df_postgresql_credit.empty or '日期' not in df_postgresql_credit.columns:
            df_postgresql_credit = pd.DataFrame({'日期': [target_dt], '全量授信申请数': [None], '全量授信申请金额': [None], '最终授信通过数': [None], '最终授信通过金额': [None], '全流程授信通过率_笔数': [None], '全流程授信通过率_金额': [None]})
            
        # 以日期为基准横向合并，使用 outer join 确保哪怕单边有数据也能保留
        df_credit = pd.merge(df_mysql_credit, df_postgresql_credit, on='日期', how='outer')
        
        # Mock 模式补齐缺失列
        if MOCK_MODE:
            required_cols = [
                '全量授信申请数', '决策授信申请数', '全量授信申请金额', '决策授信申请金额',
                '决策授信通过数', '决策授信通过金额', '最终授信通过数', '最终授信通过金额',
                '全流程授信通过率_笔数', '全流程授信通过率_金额'
            ]
            for col in required_cols:
                if col not in df_credit.columns:
                    df_credit[col] = None
        
        # 【3】增加决策授信通过率计算，防止分母为0导致报错
        df_credit['决策授信通过率_笔数'] = df_credit['决策授信通过数'] / df_credit['决策授信申请数'].replace(0, np.nan)
        df_credit['决策授信通过率_金额'] = df_credit['决策授信通过金额'] / df_credit['决策授信申请金额'].replace(0, np.nan)
        # 将 NaN 统一转换为 None 以便写入 Excel 时更安全
        df_credit['决策授信通过率_笔数'] = df_credit['决策授信通过率_笔数'].where(pd.notnull(df_credit['决策授信通过率_笔数']), None)
        df_credit['决策授信通过率_金额'] = df_credit['决策授信通过率_金额'].where(pd.notnull(df_credit['决策授信通过率_金额']), None)
        # 显式锁定授信进件表的列顺序 (根据用户需求，全量字段放到最前面)
        desired_columns_credit = [
            '日期', '全量授信申请数', '全量授信申请金额', 
            '决策授信申请数', '决策授信申请金额', '决策授信通过数', '决策授信通过金额',
            '最终授信通过数', '最终授信通过金额',
            '决策授信通过率_笔数', '决策授信通过率_金额',
            '全流程授信通过率_笔数', '全流程授信通过率_金额'
        ]
        df_credit = df_credit.reindex(columns=desired_columns_credit)
    else:
        df_credit = pd.DataFrame()
    
    # ----------------------------------------------------
    # 表 2. 客户分层 (纯 MySQL 数据)
    # ----------------------------------------------------
    sql_mysql_tier = MySQL客户分层
    df_tier = fetch_data(sql_mysql_tier, mysql_conn, target_date)
    
    # 显式锁定客户分层表的列顺序
    if not df_tier.empty:
        desired_columns_tier = [
            '客户等级', '决策授信申请数', '申请数占比', '决策授信申请金额',
            '授信通过笔数', '通过数占比', '授信通过金额', '通过金额占比',
            '授信通过率_笔数', '授信通过率_金额', '授信申请户均', '授信通过户均'
        ]
        df_tier = df_tier.reindex(columns=desired_columns_tier)
    
    # ----------------------------------------------------
    # 表 3-5. 客户分seg分析 (纯 MySQL 数据)
    # ----------------------------------------------------
    sql_seg_1 = SQL_1_SEG
    df_seg_1 = fetch_data(sql_seg_1, mysql_conn, target_date)
    
    sql_seg_2 = SQL_2_HEADCOUNT
    df_seg_2 = fetch_data(sql_seg_2, mysql_conn, target_date)
    
    sql_seg_3 = SQL_3_AMOUNT
    df_seg_3 = fetch_data(sql_seg_3, mysql_conn, target_date)
    
    # ----------------------------------------------------
    # 表 6. 支用进件_新客&白名单 (由 MySQL 和 PostgreSQL 数据垂直拼接)
    # ----------------------------------------------------
    sql_mysql_disburse_detail = MySQL支用进件明细
    df_mysql_disburse_detail = fetch_data(sql_mysql_disburse_detail, mysql_conn, target_date)
    
    # 拉取 PG 端的支用金额字典
    sql_pg_amt = PG支用金额字典
    df_pg_amt = fetch_data(sql_pg_amt, postgresql_conn, target_date)
    
    # 内存合并与计算 MySQL 端支用统计
    if not df_mysql_disburse_detail.empty and not df_pg_amt.empty:
        # MySQL 端本来就是字符串，只需去空格
        df_mysql_disburse_detail['base_id'] = df_mysql_disburse_detail['base_id'].astype(str).str.strip()
        
        # PG 端因为长数字被转成了科学计数法浮点数，我们需要强制转为数字，再用不带小数的格式化字符串输出
        # errors='coerce' 保证即使有非数字字符也不会报错
        
        
        df_mysql_merged = pd.merge(df_mysql_disburse_detail, df_pg_amt, on='base_id', how='left')
        df_mysql_merged['real_loan_amount'] = df_mysql_merged['real_loan_amount'].fillna(0).astype('float64')
        
        # 统计缺失金额的笔数并警告
        missing_amt = (df_mysql_merged['real_loan_amount'] == 0).sum()
        if missing_amt > 0:
            print(f"[警告] {target_date} 有 {missing_amt} 笔 MySQL 支用进件在 PG 中匹配不到 dn_amt！")
            # --- 增加临时诊断代码：打印前 5 个匹配不上的 base_id，让我们看看格式到底差在哪里 ---
            sample_missing = df_mysql_merged[df_mysql_merged['real_loan_amount'] == 0]['base_id'].head(5).tolist()
            print(f"[诊断] 匹配失败的 MySQL base_id 样本: {sample_missing}")
            if not df_pg_amt.empty:
                print(f"[诊断] PG 字典里的 base_id 样本: {df_pg_amt['base_id'].head(5).tolist()}")
            # -----------------------------------------------------------------------------------
            
        mysql_apply_count = len(df_mysql_merged)
        mysql_apply_amt = df_mysql_merged['real_loan_amount'].sum()
        
        df_mysql_passed = df_mysql_merged[df_mysql_merged['admit'] == '1']
        mysql_pass_count = len(df_mysql_passed)
        mysql_pass_amt = df_mysql_passed['real_loan_amount'].sum()
        
        df_mysql_disburse = pd.DataFrame({
            '日期': [target_date],
            '决策支用申请数': [mysql_apply_count],
            '决策支用申请金额': [mysql_apply_amt],
            '决策支用通过数': [mysql_pass_count],
            '决策支用通过金额': [mysql_pass_amt]
        })
    else:
        df_mysql_disburse = pd.DataFrame({'日期': [target_date], '决策支用申请数': [None], '决策支用申请金额': [None], '决策支用通过数': [None], '决策支用通过金额': [None]})
    
    
    sql_postgresql_disburse = Oracle支用进件.replace('{dynamic_cust_id_condition}', dynamic_c_cust_id_cond)
    df_postgresql_disburse = fetch_data(sql_postgresql_disburse, postgresql_conn, target_date)
    
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
        
        # 融合加权平均定价 (SQL_4_WEIGHTED_DETAIL)
        sql_weighted_detail = SQL_4_WEIGHTED_DETAIL
        df_weighted_detail = fetch_data(sql_weighted_detail, mysql_conn, target_date)
        
        avg_apply_rate = None
        avg_pass_rate = None
        
        if not df_weighted_detail.empty and not df_pg_amt.empty:
            # 同样强制转换 base_id 为字符串类型，反科学计数法
            df_weighted_detail['base_id'] = df_weighted_detail['base_id'].astype(str).str.strip()
            df_pg_amt['base_id'] = pd.to_numeric(df_pg_amt['base_id'], errors='coerce').apply(
                lambda x: f"{x:.0f}" if pd.notnull(x) else str(x)
            ).str.strip()
            
            df_weighted_merged = pd.merge(df_weighted_detail, df_pg_amt, on='base_id', how='left')
            df_weighted_merged['real_loan_amount'] = df_weighted_merged['real_loan_amount'].fillna(0).astype('float64')
            df_weighted_merged['rate'] = df_weighted_merged['rate'].fillna(0).astype('float64')
            
            # 申请阶段加权
            total_amt = df_weighted_merged['real_loan_amount'].sum()
            if total_amt > 0:
                weighted_sum = (df_weighted_merged['real_loan_amount'] * df_weighted_merged['rate']).sum()
                avg_apply_rate = f"{round(weighted_sum / total_amt, 2)}%"
                
            # 通过阶段加权（双重校验：MySQL 的 admit = 1 且 PG 的 wf_appr_sts = 997）
            df_weighted_merged['is_pass_mysql'] = df_weighted_merged['admit'].astype(str) == '1'
            if 'is_passed_pg' in df_weighted_merged.columns:
                df_weighted_merged['is_pass'] = df_weighted_merged['is_pass_mysql'] & (df_weighted_merged['is_passed_pg'] == 1)
            else:
                df_weighted_merged['is_pass'] = df_weighted_merged['is_pass_mysql']
                
            df_weighted_passed = df_weighted_merged[df_weighted_merged['is_pass']]
            total_pass_amt = df_weighted_passed['real_loan_amount'].sum()
            if total_pass_amt > 0:
                weighted_pass_sum = (df_weighted_passed['real_loan_amount'] * df_weighted_passed['rate']).sum()
                avg_pass_rate = f"{round(weighted_pass_sum / total_pass_amt, 2)}%"
        
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
        
        df_disburse['决策支用通过率_笔数'] = df_disburse['决策支用通过数'] / df_disburse['决策支用申请数'].replace(0, np.nan)
        df_disburse['决策支用通过率_金额'] = df_disburse['决策支用通过金额'] / df_disburse['决策支用申请金额'].replace(0, np.nan)
        
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
    # 表 7. 余额统计 (纯 PostgreSQL 数据)
    # ----------------------------------------------------
    sql_postgresql_balance = Oracle余额统计.replace('{dynamic_cust_id_condition}', dynamic_c_cust_id_cond)
    df_balance = fetch_data(sql_postgresql_balance, postgresql_conn, target_date, is_balance=True)
    
    # ----------------------------------------------------
    # 表 8. 累计放款 (纯 PostgreSQL 数据)
    # ----------------------------------------------------
    sql_postgresql_cum_loan = Oracle月累计放款额.replace('{dynamic_cust_id_condition}', dynamic_cust_id_cond)
    df_cum_loan = fetch_data(sql_postgresql_cum_loan, postgresql_conn, target_date, start_month=start_month, end_month=end_month)

    # ----------------------------------------------------
    # 表 9. 余额统计_星选全量 (纯 PostgreSQL 数据)
    # ----------------------------------------------------
    sql_postgresql_balance_full = Oracle余额统计_星选全量
    # 全量统计不需要替换白名单条件，直接查
    df_balance_full = fetch_data(sql_postgresql_balance_full, postgresql_conn, target_date, is_balance=True)

    # ----------------------------------------------------
    # 表 10. 累计放款_星选全量 (纯 PostgreSQL 数据)
    # ----------------------------------------------------
    sql_postgresql_cum_loan_full = Oracle月累计放款额_星选全量
    df_cum_loan_full = fetch_data(sql_postgresql_cum_loan_full, postgresql_conn, target_date, start_month=start_month,
                                  end_month=end_month)

    # ----------------------------------------------------
    # 表 11. 提价监控 (纯 PostgreSQL 数据)
    # ----------------------------------------------------
    # 提价监控的底池表中没有 c 的别名，直接使用 cust_id
    dynamic_cust_id_cond_no_alias = dynamic_c_cust_id_cond.replace('c.cust_id', 'cust_id')
    sql_price_monitor = PGSQL_PRICE_MONITOR.replace('{dynamic_cust_id_condition_no_alias}', dynamic_cust_id_cond_no_alias)
    df_price_monitor = fetch_data(sql_price_monitor, postgresql_conn, target_date)
    
    if not df_price_monitor.empty:
        # 提取指标用于控制台打印
        cnt_base = df_price_monitor['基础授信底池人数'].iloc[0]
        cnt_pricing = df_price_monitor['过滤定价小于23.58人数'].iloc[0]
        cnt_denom = df_price_monitor['分母总人数'].iloc[0]
        cnt_num_8 = df_price_monitor['90天应提未提客户数（首支超8天）'].iloc[0]
        ratio_8 = df_price_monitor['90天应提未提客户数占比（首支超8天）'].iloc[0]
        cnt_num_90 = df_price_monitor['90天应提未提客户数（首支超90天）'].iloc[0]
        ratio_90 = df_price_monitor['90天应提未提客户数占比（首支超90天）'].iloc[0]

        print(f"\n{'='*50}")
        print(f" [{target_date}] 提价监控数据核查")
        print(f"{'='*50}")
        print(f"1. 基础授信底池人数(含白名单) : {cnt_base} 人")
        print(f"2. 过滤定价<23.58%后人数      : {cnt_pricing} 人")
        print(f"3. 过滤首支且超8天(最终分母)   : {cnt_denom} 人")
        print(f"4. 90天应提未提客户数(超8天)   : {cnt_num_8} 人")
        print(f"5. 90天应提未提客户数占比(超8天): {ratio_8}")
        print(f"6. 90天应提未提客户数(超90天)  : {cnt_num_90} 人")
        print(f"7. 90天应提未提客户数占比(超90天): {ratio_90}")
        print(f"{'='*50}\n")
        
        # 打印完后删掉辅助列，以免污染最终的 Excel 输出表结构
        df_price_monitor = df_price_monitor.drop(columns=[
            '基础授信底池人数', '过滤定价小于23.58人数', '分母总人数'
        ])

    # ==========================================
    # 严格按照业务要求组装 11 张表的输出顺序
    # ==========================================
    tables_sequence = [
        ("授信进件", df_credit),
        ("客户分层", df_tier),
        ("客户分seg", df_seg_1),
        ("客户分层&分seg_人头", df_seg_2),
        ("客户分层&分seg_金额", df_seg_3),
        ("支用进件_新客&白名单", df_disburse),
        ("余额统计_嵩海", df_balance),
        ("累计放款_嵩海", df_cum_loan),
        ("余额统计_星选全量", df_balance_full),
        ("累计放款_星选全量", df_cum_loan_full),
        ("提价监控", df_price_monitor)
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
        title_cell.font = title_font
        
        # 如果 df 是空的（数据为空且所有列全为 None/NaN，或者没列），为了防止 openpyxl 后续报错，我们需要跳过后续写表头和数据的过程
        if df.empty or df.dropna(how='all').empty:
             current_row += 2
             continue
             
        # [动作 2]：写入表头 (DataFrame 的列名)
        header_row = current_row + 1
        for col_idx, col_name in enumerate(df.columns, start=1):
            cell = ws.cell(row=header_row, column=col_idx, value=col_name)
            cell.alignment = Alignment(horizontal='center')
            
        # [动作 3]：逐行写入业务数据
        data_start_row = header_row + 1
        data_rows_count = 0  
        
        if not df.empty:
            for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=False), start=data_start_row):
                for c_idx, value in enumerate(row, start=1):
                    ws.cell(row=r_idx, column=c_idx, value=value)
                data_rows_count += 1
                    
        # [动作 4]：计算自适应列宽
        for col_idx, col_name in enumerate(df.columns, start=1):
            column_letter = get_column_letter(col_idx)
            current_width = ws.column_dimensions[column_letter].width or 0
            
            max_len = len(str(col_name).encode('gbk', errors='ignore'))
            if not df.empty:
                for item in df[col_name]:
                    item_len = len(str(item).encode('gbk', errors='ignore'))
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
    
    # 4. 组装最终的输出文件绝对路径
    output_filename = os.path.join(output_dir, f"晋商项目报表_全量处理结果_{start_str.replace('-', '')}_至_{end_str.replace('-', '')}.xlsx")
    
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
                    writer.book.active = 0
                    del writer.book["Sheet"]
                elif "Sheet" in writer.book.sheetnames and len(writer.book.sheetnames) == 1:
                    ws = writer.book["Sheet"]
                    ws.cell(row=1, column=1, value="无数据生成")
                
        print(f"\n[成功] 所有日期数据处理完毕，报表已一次性保存至: {output_filename}")
        
    except Exception as e:
        print(f"\n[严重错误] 执行期间发生异常: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
