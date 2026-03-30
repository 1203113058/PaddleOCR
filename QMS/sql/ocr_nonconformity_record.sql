-- 不合格识别记录表（PPStructure 力学/表格判定）
-- 数据库：sg_local   字符集：utf8mb4

CREATE DATABASE IF NOT EXISTS `sg_local`
  DEFAULT CHARACTER SET utf8mb4
  DEFAULT COLLATE utf8mb4_unicode_ci;

USE `sg_local`;

CREATE TABLE IF NOT EXISTS `ocr_nonconformity_record` (
  `id`              BIGINT UNSIGNED NOT NULL AUTO_INCREMENT COMMENT '主键',
  `report_name`     VARCHAR(512)    NOT NULL DEFAULT '' COMMENT '报告文件名（不含路径）',
  `report_type`     VARCHAR(64)     NOT NULL DEFAULT '' COMMENT '报告类型：力学检测等',
  `input_full_path` VARCHAR(1024)   NOT NULL DEFAULT '' COMMENT '原始 PDF/图片完整路径',
  `sheet_name`      VARCHAR(128)    NOT NULL DEFAULT '' COMMENT 'Excel 工作表名',
  `page_no`         INT             NULL COMMENT '来源 PDF 页码',
  `table_index`     INT             NULL COMMENT '该页内表格序号',
  `sample_no`       VARCHAR(128)    NOT NULL DEFAULT '' COMMENT '试样编号',
  `test_item`       VARCHAR(768)    NOT NULL DEFAULT '' COMMENT '检测项目（表头推断，如屈服强度）',
  `standard_text`   VARCHAR(512)    NOT NULL DEFAULT '' COMMENT '标准/要求值（单元格原文）',
  `rule_type`       VARCHAR(16)     NOT NULL DEFAULT '' COMMENT '判定规则：range/gte/lte/eq',
  `actual_value`    VARCHAR(512)    NOT NULL DEFAULT '' COMMENT '实测值',
  `fail_reason`     VARCHAR(1024)   NOT NULL DEFAULT '' COMMENT '不合格原因',
  `excel_row`       INT             NULL COMMENT 'Excel 行号',
  `excel_col`       INT             NULL COMMENT 'Excel 列号',
  `excel_col_letter` VARCHAR(8)     NOT NULL DEFAULT '' COMMENT 'Excel 列字母',
  `xlsx_path`       VARCHAR(1024)   NOT NULL DEFAULT '' COMMENT '输出 Excel 路径',
  `batch_id`        VARCHAR(64)     NOT NULL DEFAULT '' COMMENT '同次识别批次（时间戳）',
  `created_at`      DATETIME        NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT '创建时间',
  `updated_at`      DATETIME        NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',
  PRIMARY KEY (`id`),
  KEY `idx_report_batch` (`batch_id`),
  KEY `idx_report_name` (`report_name`(191)),
  KEY `idx_created` (`created_at`),
  KEY `idx_sample` (`sample_no`(64))
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
  COMMENT='OCR 表格识别-不合格项记录';
