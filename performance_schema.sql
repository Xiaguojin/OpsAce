-- ============================================================
-- 绩效管理系统 - 数据库表结构设计
-- 支持: 二级部门-三级部门层级联动、Excel解析存储、自关联血缘追溯
-- Database: PostgreSQL 14+
-- ============================================================

-- ========== 1. 部门表 (自关联树形结构) ==========
CREATE TABLE department (
    id              UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    dept_code       VARCHAR(50)  NOT NULL UNIQUE,      -- 部门编码, 如 "SW-ENG-L2"
    dept_name       VARCHAR(100) NOT NULL,              -- 部门名称, 如 "集成维护开发部"
    dept_level      SMALLINT     NOT NULL,              -- 层级: 2=二级部门, 3=三级部门
    parent_dept_id  UUID REFERENCES department(id),     -- 自关联: 三级部门指向二级部门
    sort_order      INT          DEFAULT 0,             -- 排序序号
    is_active       BOOLEAN      DEFAULT TRUE,
    created_at      TIMESTAMPTZ  DEFAULT NOW(),
    updated_at      TIMESTAMPTZ  DEFAULT NOW()
);

-- 层级查询索引
CREATE INDEX idx_dept_parent ON department(parent_dept_id);
CREATE INDEX idx_dept_level  ON department(dept_level);

-- ========== 2. KPI 关键绩效指标表 (自关联) ==========
CREATE TABLE performance_kpi (
    id                  UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    
    -- 部门关联
    dept_id             UUID NOT NULL REFERENCES department(id),
    
    -- 自关联: 三级部门的KPI指向二级部门的KPI (parent_kpi_id = 二级KPI.id)
    -- 二级部门的KPI: parent_kpi_id = NULL
    -- 三级部门的KPI: parent_kpi_id = 对应的二级部门KPI.id
    parent_kpi_id       UUID REFERENCES performance_kpi(id),
    
    -- Excel原始数据字段
    indicator_name      VARCHAR(200)  NOT NULL,         -- 指标名称
    definition_desc     TEXT,                            -- 定义描述
    category            VARCHAR(100),                    -- 所属维度/分类 (如: 软件交付质量)
    weight              DECIMAL(5,2)  NOT NULL DEFAULT 0,-- 权重 (0.00-1.00, 即0%-100%)
    unit                VARCHAR(20),                     -- 单位 (如: %, 个)
    data_source         VARCHAR(200),                    -- 数据来源/数据提供部门
    last_year_value     VARCHAR(100),                    -- 上周期完成值
    start_date          DATE,                            -- 开始日期
    end_date            DATE,                            -- 结束日期
    
    -- 目标值 (三档: 门槛值60分 / 目标值100分 / 挑战值140分)
    threshold_value     VARCHAR(100),                    -- 门槛值 (60分线)
    target_value        VARCHAR(100),                    -- 目标值 (100分线)
    challenge_value     VARCHAR(100),                    -- 挑战值 (140分线)
    
    -- 周期目标值 (H1/年度 分别存储)
    h1_threshold        VARCHAR(100),                    -- H1门槛值
    h1_target           VARCHAR(100),                    -- H1目标值
    h1_challenge        VARCHAR(100),                    -- H1挑战值
    annual_threshold    VARCHAR(100),                    -- 年度门槛值
    annual_target       VARCHAR(100),                    -- 年度目标值
    annual_challenge    VARCHAR(100),                    -- 年度挑战值
    
    -- Excel解析追溯
    source_type         VARCHAR(20) DEFAULT 'excel',     -- 数据来源: excel / manual
    excel_row_ref       INT,                             -- Excel原始行号 (用于覆盖更新时定位)
    excel_file_name     VARCHAR(255),                    -- 来源Excel文件名
    fiscal_year         INT NOT NULL,                    -- 财年 (如: 2026)
    
    -- 状态与进度
    status              VARCHAR(20) DEFAULT 'active',    -- active / archived / draft
    current_value       VARCHAR(100),                    -- 当前完成值
    completion_rate     DECIMAL(5,2) DEFAULT 0,          -- 完成率 (0-100%)
    
    created_at          TIMESTAMPTZ DEFAULT NOW(),
    updated_at          TIMESTAMPTZ DEFAULT NOW(),
    
    -- 约束: 同一部门同一财年同一指标名不可重复
    CONSTRAINT uk_kpi_dept_year_name UNIQUE (dept_id, fiscal_year, indicator_name)
);

-- 关键查询索引
CREATE INDEX idx_kpi_dept        ON performance_kpi(dept_id);
CREATE INDEX idx_kpi_parent      ON performance_kpi(parent_kpi_id);
CREATE INDEX idx_kpi_fiscal_year ON performance_kpi(fiscal_year);
CREATE INDEX idx_kpi_category    ON performance_kpi(category);

-- ========== 3. 重点工作表 (自关联) ==========
CREATE TABLE performance_task (
    id                  UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    
    -- 部门关联
    dept_id             UUID NOT NULL REFERENCES department(id),
    
    -- 自关联: 三级部门的重点工作指向二级部门的重点工作
    parent_task_id      UUID REFERENCES performance_task(id),
    
    -- 基本信息
    task_name           VARCHAR(200) NOT NULL,           -- 任务名称
    task_desc           TEXT,                             -- 任务描述
    category            VARCHAR(100),                     -- 所属维度/分类
    weight              DECIMAL(5,2) NOT NULL DEFAULT 0, -- 权重
    start_date          DATE,                             -- 开始日期
    end_date            DATE,                             -- 结束日期
    
    -- 承接模式: 三级部门可以 "完整承接" 或 "部分承接(拆解)"
   承接模式             VARCHAR(20) DEFAULT 'full',      -- full=完整承接, partial=部分承接
    
    -- Excel解析追溯
    source_type         VARCHAR(20) DEFAULT 'excel',
    excel_row_ref       INT,                              -- 重点工作名称所在行号
    excel_merge_range   VARCHAR(50),                      -- 合并单元格范围 (如 "B5:B7")
    excel_file_name     VARCHAR(255),
    fiscal_year         INT NOT NULL,
    
    -- 状态与进度
    status              VARCHAR(20) DEFAULT 'active',
    current_progress    DECIMAL(5,2) DEFAULT 0,           -- 当前进度 (0-100%)
    
    created_at          TIMESTAMPTZ DEFAULT NOW(),
    updated_at          TIMESTAMPTZ DEFAULT NOW(),
    
    CONSTRAINT uk_task_dept_year_name UNIQUE (dept_id, fiscal_year, task_name)
);

CREATE INDEX idx_task_dept        ON performance_task(dept_id);
CREATE INDEX idx_task_parent      ON performance_task(parent_task_id);
CREATE INDEX idx_task_fiscal_year ON performance_task(fiscal_year);

-- ========== 4. 里程碑表 (H1/H2) ==========
CREATE TABLE performance_milestone (
    id                  UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    task_id             UUID NOT NULL REFERENCES performance_task(id) ON DELETE CASCADE,
    
    milestone_type      VARCHAR(5) NOT NULL,              -- H1 / H2
    milestone_name      VARCHAR(300) NOT NULL,            -- 里程碑名称/描述
    key_tasks           JSONB,                            -- 关键任务列表 (JSON数组)
    
    start_date          DATE,
    end_date            DATE,
    
    -- Excel解析追溯
    excel_row_ref       INT,
    
    -- 状态
    status              VARCHAR(20) DEFAULT 'pending',    -- pending / in_progress / done / blocked
    completion_rate     DECIMAL(5,2) DEFAULT 0,
    
    created_at          TIMESTAMPTZ DEFAULT NOW(),
    updated_at          TIMESTAMPTZ DEFAULT NOW(),
    
    CONSTRAINT uk_milestone_task_type UNIQUE (task_id, milestone_type)
);

CREATE INDEX idx_milestone_task ON performance_milestone(task_id);

-- ========== 5. 加减分项表 ==========
CREATE TABLE performance_bonus (
    id                  UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    dept_id             UUID NOT NULL REFERENCES department(id),
    fiscal_year         INT NOT NULL,
    
    title               VARCHAR(200) NOT NULL,
    description         TEXT,
    rule                TEXT,                             -- 加减分规则描述
    threshold_value     VARCHAR(100),                     -- 门槛值
    target_value        VARCHAR(100),                     -- 目标值
    challenge_value     VARCHAR(100),                     -- 挑战值
    
    score_delta         DECIMAL(5,2) DEFAULT 0,           -- 实际加减分
    source_type         VARCHAR(20) DEFAULT 'excel',
    excel_row_ref       INT,
    
    created_at          TIMESTAMPTZ DEFAULT NOW(),
    updated_at          TIMESTAMPTZ DEFAULT NOW()
);

CREATE INDEX idx_bonus_dept ON performance_bonus(dept_id);

-- ========== 6. Excel上传批次记录 (支持覆盖/增量更新) ==========
CREATE TABLE excel_upload_batch (
    id                  UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    dept_id             UUID NOT NULL REFERENCES department(id),
    fiscal_year         INT NOT NULL,
    
    file_name           VARCHAR(255) NOT NULL,
    file_hash           VARCHAR(64),                      -- 文件MD5, 用于去重
    upload_mode         VARCHAR(20) NOT NULL DEFAULT 'overwrite', -- overwrite / incremental
    upload_user         VARCHAR(100),
    
    -- 解析结果
    total_rows          INT DEFAULT 0,
    kpi_count           INT DEFAULT 0,
    task_count          INT DEFAULT 0,
    bonus_count         INT DEFAULT 0,
    weight_total        DECIMAL(5,2),                     -- 权重合计 (应=1.00)
    weight_valid        BOOLEAN DEFAULT FALSE,            -- 权重是否=100%
    parse_errors        JSONB,                            -- 解析错误列表
    
    status              VARCHAR(20) DEFAULT 'processing', -- processing / success / failed
    created_at          TIMESTAMPTZ DEFAULT NOW()
);

CREATE INDEX idx_batch_dept ON excel_upload_batch(dept_id);

-- ============================================================
-- 视图: 二级部门全景 - 查看所有三级部门承接情况
-- ============================================================
CREATE OR REPLACE VIEW v_l2_kpi_acceptance AS
SELECT
    k2.id               AS l2_kpi_id,
    k2.indicator_name   AS l2_kpi_name,
    k2.category,
    k2.weight           AS l2_weight,
    k2.target_value     AS l2_target,
    k3.id               AS l3_kpi_id,
    d3.dept_name        AS l3_dept_name,
    k3.target_value     AS l3_target,
    k3.completion_rate  AS l3_completion,
    k3.status           AS l3_status
FROM performance_kpi k2
JOIN department d2 ON k2.dept_id = d2.id AND d2.dept_level = 2
LEFT JOIN performance_kpi k3 ON k3.parent_kpi_id = k2.id
LEFT JOIN department d3 ON k3.dept_id = d3.id
WHERE k2.parent_kpi_id IS NULL;

-- 同理: 重点工作承接视图
CREATE OR REPLACE VIEW v_l2_task_acceptance AS
SELECT
    t2.id               AS l2_task_id,
    t2.task_name        AS l2_task_name,
    t2.category,
    t2.weight           AS l2_weight,
    t3.id               AS l3_task_id,
    d3.dept_name        AS l3_dept_name,
    t3.current_progress AS l3_progress,
    t3.status           AS l3_status,
    m.milestone_type,
    m.milestone_name,
    m.completion_rate   AS milestone_completion,
    m.status            AS milestone_status
FROM performance_task t2
JOIN department d2 ON t2.dept_id = d2.id AND d2.dept_level = 2
LEFT JOIN performance_task t3 ON t3.parent_task_id = t2.id
LEFT JOIN department d3 ON t3.dept_id = d3.id
LEFT JOIN performance_milestone m ON m.task_id = t3.id
WHERE t2.parent_task_id IS NULL
ORDER BY t2.sort_order, t3.dept_name, m.milestone_type;

-- ============================================================
-- 权重校验函数: 检查某部门某年KPI+重点工作权重是否=100%
-- ============================================================
CREATE OR REPLACE FUNCTION check_weight_total(
    p_dept_id UUID,
    p_fiscal_year INT
) RETURNS TABLE(section TEXT, total_weight DECIMAL(5,2), is_valid BOOLEAN) AS $$
BEGIN
    RETURN QUERY
    SELECT 'KPI'::TEXT,
           COALESCE(SUM(weight), 0),
           COALESCE(SUM(weight), 0) = 1.00
    FROM performance_kpi
    WHERE dept_id = p_dept_id AND fiscal_year = p_fiscal_year AND parent_kpi_id IS NULL
    
    UNION ALL
    
    SELECT '重点工作'::TEXT,
           COALESCE(SUM(weight), 0),
           COALESCE(SUM(weight), 0) = 1.00
    FROM performance_task
    WHERE dept_id = p_dept_id AND fiscal_year = p_fiscal_year AND parent_task_id IS NULL
    
    UNION ALL
    
    SELECT '合计'::TEXT,
           COALESCE((SELECT SUM(weight) FROM performance_kpi WHERE dept_id = p_dept_id AND fiscal_year = p_fiscal_year AND parent_kpi_id IS NULL), 0)
           + COALESCE((SELECT SUM(weight) FROM performance_task WHERE dept_id = p_dept_id AND fiscal_year = p_fiscal_year AND parent_task_id IS NULL), 0),
           COALESCE((SELECT SUM(weight) FROM performance_kpi WHERE dept_id = p_dept_id AND fiscal_year = p_fiscal_year AND parent_kpi_id IS NULL), 0)
           + COALESCE((SELECT SUM(weight) FROM performance_task WHERE dept_id = p_dept_id AND fiscal_year = p_fiscal_year AND parent_task_id IS NULL), 0) = 1.00;
END;
$$ LANGUAGE plpgsql;
