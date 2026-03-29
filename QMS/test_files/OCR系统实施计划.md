# OCR 检测报告识别系统 — 实施计划

**版本：** V1.0  
**日期：** 2026年3月7日  
**部署方式：** 完全本地部署（内网，数据不出网）  
**核心框架：** PaddleOCR（本地推理 + 本地训练）

---

## 一、项目概述

### 1.1 目标

基于 PaddleOCR 在本地/内网服务器上搭建一套 OCR 自动识别系统，实现对铸件、锻件等质量检测报告（PDF扫描件）的关键字段自动提取，输出结构化数据（JSON/Excel），并标记低置信度字段供人工复核。**本阶段不涉及 QMS 系统对接，专注识别准确性和标注训练能力。**

### 1.2 报告类型（当前已有样本）

| 报告类型 | 文件名 | 关键字段 |
|---------|--------|---------|
| 铸件质量证明书（多页） | 铸件质量证明书1（多页报告）.pdf | 化学成分、力学性能、基本信息 |
| 锻件质量证明书 | 锻件质量证明书.pdf | 化学成分、力学性能、基本信息 |
| 化学成分复检报告 | 化学成分复检报告.pdf | 各元素实测值 |
| 机械性能复检报告 | 机械性能复检报告.pdf | 抗拉强度、屈服强度、延伸率、硬度等 |

### 1.3 技术架构

```
PDF文件输入
    ↓
[预处理层]  pdf2image → 去噪 / 纠偏 / 对比度增强（OpenCV）
    ↓
[检测层]    PP-OCRv4 文本检测（同时检测印章遮挡区域）
    ↓
[识别层]    PP-OCRv4 中英文混合识别
    ↓
[结构层]    PP-Structure 表格结构解析
    ↓
[理解层]    KIE（关键信息提取）按模板定位字段
    ↓
[校验层]    置信度过滤 + 印章遮挡标记 + 数值合规判断
    ↓
[输出层]    结构化 JSON / Excel（字段 + 置信度 + 复核标记）
```

---

## 二、硬件要求

> **首要确认项：运行前请先检查服务器配置**

```bash
nvidia-smi        # 查看 GPU 情况
free -h           # 查看内存
python3 --version # 确认 Python 版本（需 3.8-3.11）
```

| 硬件条件 | 推理速度 | 训练可行性 | 推荐程度 |
|---------|---------|----------|---------|
| NVIDIA GPU ≥8GB（Linux服务器） | 0.5-2秒/页 | 本地训练，效率高 | ★★★★★ 最推荐 |
| 仅CPU + ≥16GB内存（Linux/Mac） | 15-30秒/页 | 可训练，速度慢 | ★★★☆☆ 开发验证用 |
| Mac M系列芯片（当前开发机） | 5-15秒/页 | 可训练，速度慢 | ★★★☆☆ 开发验证用 |

---

## 三、阶段划分与里程碑

```
阶段一（Week 1-2）：环境搭建 + 基础识别验证
    ↓ 里程碑：4类报告可跑通完整识别链路，输出 JSON
阶段二（Week 3-4）：文档分析 + 字段规则定义
    ↓ 里程碑：定义完所有字段提取规则，建立模板配置文件
阶段三（Week 5-7）：PPOCRLabel 标注 + KIE 模型训练
    ↓ 里程碑：铸件/锻件报告字段准确率 ≥ 95%
阶段四（Week 8）：  印章检测集成 + 置信度优化
    ↓ 里程碑：自动标记「需复核」字段，准确率 ≥ 98%
阶段五（Week 9-10）：压力测试 + 输出格式验收
    ↓ 里程碑：可交付质检人员试用，支持 200份/天
```

---

## 四、阶段一：环境搭建与基础识别验证

### 4.1 创建虚拟环境

```bash
conda create -n ocr_local python=3.10
conda activate ocr_local
```

### 4.2 安装 PaddlePaddle（根据硬件选择）

```bash
# === 有 NVIDIA GPU（CUDA 11.8）===
pip install paddlepaddle-gpu==2.6.1.post118 \
  -f https://www.paddlepaddle.org.cn/whl/linux/mkl/avx/stable.html

# === 有 NVIDIA GPU（CUDA 12.3）===
pip install paddlepaddle-gpu==2.6.1.post120 \
  -f https://www.paddlepaddle.org.cn/whl/linux/mkl/avx/stable.html

# === 仅 CPU / Mac ===
pip install paddlepaddle==2.6.1
```

### 4.3 安装 PaddleOCR 及文档处理依赖

```bash
pip install paddleocr>=2.7.0
pip install pdf2image pillow opencv-python-headless numpy pandas openpyxl

# macOS 安装 poppler（pdf2image 底层依赖）
brew install poppler

# Linux 安装 poppler
# apt-get install poppler-utils     # Ubuntu/Debian
# yum install poppler-utils         # CentOS/RHEL
```

### 4.4 验证安装

```python
import paddle
import paddleocr
print("PaddlePaddle 版本:", paddle.__version__)
print("GPU 是否可用:", paddle.is_compiled_with_cuda())
print("PaddleOCR 安装正常")
```

### 4.5 基础冒烟测试脚本

> 新建 `/Users/project/OCR/first_run.py`

```python
from paddleocr import PaddleOCR
from pdf2image import convert_from_path
import json, os

# 首次运行会自动下载预训练模型（约500MB），需保持网络畅通
ocr = PaddleOCR(
    use_angle_cls=True,  # 启用方向分类（处理旋转扫描件）
    lang='ch',           # 中英文混合识别
    use_gpu=False,       # 有GPU改为 True
    show_log=False,
)

def process_pdf(pdf_path, output_json_path):
    print(f"正在处理: {pdf_path}")
    images = convert_from_path(pdf_path, dpi=300)
    print(f"共 {len(images)} 页")

    all_results = []
    for page_num, img in enumerate(images):
        tmp_path = f'/tmp/ocr_page_{page_num}.jpg'
        img.save(tmp_path, 'JPEG', quality=95)

        result = ocr.ocr(tmp_path, cls=True)
        page_data = {"page": page_num + 1, "lines": []}

        if result and result[0]:
            for line in result[0]:
                text = line[1][0]
                confidence = line[1][1]
                page_data["lines"].append({
                    "text": text,
                    "confidence": round(confidence, 4),
                    "bbox": line[0],
                    "need_review": confidence < 0.90
                })

        all_results.append(page_data)
        print(f"  第{page_num+1}页: 识别 {len(page_data['lines'])} 行，"
              f"需复核 {sum(1 for l in page_data['lines'] if l['need_review'])} 行")

    os.makedirs(os.path.dirname(output_json_path), exist_ok=True)
    with open(output_json_path, 'w', encoding='utf-8') as f:
        json.dump(all_results, f, ensure_ascii=False, indent=2)
    print(f"结果已保存: {output_json_path}")
    return all_results

# 运行测试
process_pdf(
    '文件/铸件质量证明书1（多页报告）.pdf',
    'output/铸件质量证明书_识别结果.json'
)
```

```bash
cd /Users/project/OCR
mkdir -p output
python first_run.py
```

**预期输出：** 生成 `output/铸件质量证明书_识别结果.json`，包含每页所有识别文字和置信度。

---

## 五、阶段二：表格结构识别与字段规则定义

### 5.1 使用 PP-Structure 解析表格

你的报告中化学成分表、机械性能表需要 PP-Structure 来保留行列结构，否则只能得到散乱文字：

```python
from paddleocr import PPStructure
from pdf2image import convert_from_path
import numpy as np

table_engine = PPStructure(table=True, ocr=True, lang='ch', show_log=False)

def extract_tables(pdf_path):
    images = convert_from_path(pdf_path, dpi=300)
    for page_num, img in enumerate(images):
        img_array = np.array(img)
        result = table_engine(img_array)
        for region in result:
            if region['type'] == 'table':
                print(f"第{page_num+1}页 表格HTML:")
                print(region['res']['html'])
```

**表格输出示例：**
```html
<table>
  <tr><td>元素</td><td>C</td><td>Si</td><td>Mn</td><td>P</td><td>S</td></tr>
  <tr><td>实测值(%)</td><td>0.23</td><td>0.45</td><td>0.78</td><td>0.012</td><td>0.008</td></tr>
  <tr><td>要求值(%)</td><td>≤0.30</td><td>≤0.60</td><td>0.60-1.00</td><td>≤0.035</td><td>≤0.035</td></tr>
</table>
```

### 5.2 建立模板字段规则配置

为每种报告建立 JSON 配置文件，存放于 `templates/` 目录：

```json
// templates/铸件质量证明书_v1.json
{
  "template_id": "casting_cert_v1",
  "template_name": "铸件质量证明书",
  "fields": {
    "basic_info": [
      {"name": "产品代号",   "type": "text",   "confidence_threshold": 0.90},
      {"name": "图号",       "type": "text",   "confidence_threshold": 0.90},
      {"name": "零部件名称", "type": "text",   "confidence_threshold": 0.90},
      {"name": "材质",       "type": "text",   "confidence_threshold": 0.90},
      {"name": "报告编号",   "type": "text",   "confidence_threshold": 0.95},
      {"name": "检测日期",   "type": "date",   "confidence_threshold": 0.95},
      {"name": "供应商名称", "type": "text",   "confidence_threshold": 0.90}
    ],
    "chemical_composition": [
      {"name": "C%",  "type": "number", "confidence_threshold": 0.95},
      {"name": "Si%", "type": "number", "confidence_threshold": 0.95},
      {"name": "Mn%", "type": "number", "confidence_threshold": 0.95},
      {"name": "P%",  "type": "number", "confidence_threshold": 0.95},
      {"name": "S%",  "type": "number", "confidence_threshold": 0.95}
    ],
    "mechanical_properties": [
      {"name": "抗拉强度Rm/MPa",    "type": "number", "confidence_threshold": 0.95},
      {"name": "屈服强度Rp0.2/MPa", "type": "number", "confidence_threshold": 0.95},
      {"name": "延伸率A%",          "type": "number", "confidence_threshold": 0.95},
      {"name": "硬度HB",            "type": "number", "confidence_threshold": 0.95},
      {"name": "冲击值J",           "type": "number", "confidence_threshold": 0.95}
    ]
  }
}
```

---

## 六、阶段三：数据标注与 KIE 模型训练

### 6.1 安装 PPOCRLabel 标注工具

```bash
pip install PPOCRLabel

# 启动 KIE（关键信息提取）标注模式
PPOCRLabel --kie True
```

### 6.2 标注工作流程

```
Step 1：将 PDF 转为图片（300 DPI）
    ↓
Step 2：导入图片到 PPOCRLabel
    ↓
Step 3：框选文字区域（工具可自动检测，手动校正）
    ↓
Step 4：为每个框打语义标签
         KEY   类型：字段名称（如"C元素标题"）
         VALUE 类型：实际数值（如 C_VALUE = "0.23"）
    ↓
Step 5：导出标注文件到 train_data/ 目录
```

### 6.3 各报告类型标注量建议

| 报告类型 | 最低标注页数 | 标注重点 | 优先级 |
|---------|-----------|---------|-------|
| 铸件质量证明书 | 80 页 | 化学成分表 + 力学性能表 | P0 |
| 锻件质量证明书 | 60 页 | 化学成分表 + 力学性能表 | P0 |
| 化学成分复检报告 | 50 页 | 复检值 vs 初检值对比 | P1 |
| 机械性能复检报告 | 50 页 | 抗拉 / 屈服 / 延伸率 | P1 |

### 6.4 训练 KIE 模型

```bash
cd /Users/project/OCR/PaddleOCR

# 启动训练（基于 LayoutXLM，专为文档理解设计）
python tools/train.py \
  -c configs/kie/vi_layoutxlm/ser_vi_layoutxlm_xfund_zh.yml \
  -o Global.epoch_num=100 \
     Global.save_model_dir=./output/my_kie_model/ \
     Train.dataset.data_dir=./train_data/casting_cert/
```

### 6.5 导出推理模型

```bash
python tools/export_model.py \
  -c configs/kie/vi_layoutxlm/ser_vi_layoutxlm_xfund_zh.yml \
  -o Global.pretrained_model=./output/my_kie_model/best_accuracy \
     Global.save_inference_dir=./inference/casting_cert_kie/
```

---

## 七、阶段四：印章检测集成

### 7.1 印章检测原理

PaddleOCR 内置印章识别 Pipeline，可检测圆形/椭圆形印章区域。被印章覆盖的字段不强行识别，直接标记为「**需复核**」，由质检人员手动填写。

```python
from paddleocr import create_pipeline

# 使用 PP-StructureV3 的印章识别 Pipeline
pipeline = create_pipeline("seal_recognition")

def check_seal_coverage(img_path):
    output = pipeline.predict(img_path)
    seal_regions = []
    for res in output:
        seal_regions.extend(res.get('seal_det_res', []))
    return seal_regions  # 返回所有印章边界框
```

### 7.2 字段遮挡判断逻辑

```python
def is_covered_by_seal(field_bbox, seal_regions, overlap_threshold=0.3):
    """
    判断某字段框是否被印章遮挡
    overlap_threshold：重叠面积比例超过30%即判定为遮挡
    """
    for seal_box in seal_regions:
        overlap = calc_iou(field_bbox, seal_box)
        if overlap > overlap_threshold:
            return True
    return False
```

---

## 八、阶段五：完整 Pipeline 与输出格式

### 8.1 完整业务流水线

```python
class ReportOCRPipeline:
    def __init__(self, template_id):
        self.ocr = PaddleOCR(use_angle_cls=True, lang='ch', use_gpu=True)
        self.table_engine = PPStructure(table=True, ocr=True, lang='ch')
        self.template = load_template(f'templates/{template_id}.json')

    def process(self, pdf_path):
        # 1. PDF 转图片（300 DPI）
        images = pdf_to_images(pdf_path, dpi=300)

        # 2. 图像预处理（去噪、纠偏、增强对比度）
        enhanced = [preprocess(img) for img in images]

        # 3. 印章检测（标记遮挡区域）
        seal_regions = [check_seal_coverage(img) for img in enhanced]

        # 4. 表格结构识别
        tables = [self.table_engine(img) for img in enhanced]

        # 5. KIE 字段提取
        extracted = self.extract_fields(tables, seal_regions)

        # 6. 置信度过滤，生成复核队列
        return self.validate_and_flag(extracted)

    def validate_and_flag(self, extracted):
        for field_name, field_data in extracted.items():
            conf = field_data['confidence']
            field_type = self.template['fields'][field_name]['type']
            threshold = 0.95 if field_type == 'number' else 0.90

            if field_data.get('seal_covered'):
                field_data['flag'] = '印章遮挡-需复核'
            elif conf < threshold:
                field_data['flag'] = f'低置信度({conf:.0%})-需复核'
            else:
                field_data['flag'] = None
        return extracted
```

### 8.2 输出 JSON 格式

```json
{
  "report_id": "RC-2026-001",
  "template": "铸件质量证明书_v1",
  "process_time": "2026-03-07T10:30:00",
  "status": "需复核",
  "fields": {
    "产品代号":       {"value": "ZG25",  "confidence": 0.9856, "flag": null},
    "报告编号":       {"value": "QC-001","confidence": 0.9923, "flag": null},
    "C%":            {"value": "0.23",  "confidence": 0.9934, "flag": null},
    "Si%":           {"value": "0.45",  "confidence": 0.9756, "flag": null},
    "Mn%":           {"value": null,    "confidence": 0.0,    "flag": "印章遮挡-需复核"},
    "抗拉强度Rm/MPa": {"value": "485",   "confidence": 0.8812, "flag": "低置信度(88%)-需复核"}
  },
  "review_queue": ["Mn%", "抗拉强度Rm/MPa"],
  "auto_pass_count": 45,
  "review_count": 3
}
```

### 8.3 同时导出 Excel

```python
import pandas as pd

def export_to_excel(result, output_path):
    rows = []
    for field_name, data in result['fields'].items():
        rows.append({
            '字段名': field_name,
            '识别值': data['value'],
            '置信度': f"{data['confidence']:.1%}",
            '状态': data['flag'] if data['flag'] else '✓ 通过',
        })
    df = pd.DataFrame(rows)
    # 高亮需复核行（红色）
    df.to_excel(output_path, index=False)
    print(f"Excel 已导出: {output_path}")
```

---

## 九、项目文件结构

```
/Users/project/OCR/
├── PaddleOCR/              # 框架源码（已有）
│   ├── tools/              # 训练/推理脚本
│   ├── ppocr/              # OCR 核心模块
│   ├── ppstructure/        # 表格/版面分析
│   └── configs/            # 训练配置文件
│
├── 文件/                   # 待识别 PDF（已有）
│   ├── 铸件质量证明书1（多页报告）.pdf
│   ├── 锻件质量证明书.pdf
│   ├── 化学成分复检报告.pdf
│   └── 机械性能复检报告.pdf
│
├── templates/              # 报告字段规则配置
│   ├── 铸件质量证明书_v1.json
│   ├── 锻件质量证明书_v1.json
│   ├── 化学成分复检报告_v1.json
│   └── 机械性能复检报告_v1.json
│
├── train_data/             # 标注数据集（PPOCRLabel 导出）
│   ├── casting_cert/       # 铸件质量证明书标注数据
│   └── forging_cert/       # 锻件质量证明书标注数据
│
├── models/                 # 训练产出的推理模型
│   ├── casting_cert_kie/
│   └── forging_cert_kie/
│
├── output/                 # 识别结果输出
│   ├── *.json
│   └── *.xlsx
│
├── first_run.py            # 冒烟测试脚本
├── ocr_pipeline.py         # 完整业务流水线
└── requirements.txt        # 项目依赖清单
```

---

## 十、精度目标与验收标准

| 场景 | 目标准确率 | 处理策略 |
|------|----------|---------|
| 电子版PDF / 标准清晰扫描件 | ≥ 98% | 直接识别 |
| 低质量扫描件（偏斜、噪点） | ≥ 95% | 图像预处理后识别 |
| 拍照PDF（变形、光线不均） | 尽力识别 | 低置信度字段全标「需复核」 |
| 印章遮挡区域 | — | 直接标「需复核」，不强行识别 |
| 数字字段置信度阈值 | ≥ 95% | 低于阈值进复核队列 |
| 文字字段置信度阈值 | ≥ 90% | 低于阈值进复核队列 |

---

## 十一、常见问题与解决方案

| 问题现象 | 原因 | 解决方案 |
|---------|------|---------|
| 小数点被识别为逗号 | 扫描噪点干扰 | 后处理正则校正 + 提高DPI到400 |
| 表格数值串行（列内容混乱） | 列间距过小 | 使用 PP-StructureV3 表格解析 |
| 中文单位字符被截断 | 字体过小 | 图像超分辨率预处理 |
| 旋转扫描件识别差 | 纸张偏斜 | 启用 `use_angle_cls=True` + OpenCV 自动纠偏 |
| 0和8混淆、1和7混淆 | 低分辨率扫描 | 增加训练数据中含类似字符的样本 |

---

## 十二、里程碑检查清单

### Week 1-2（环境搭建 + 基础验证）
- [ ] 确认服务器 GPU / CPU 配置
- [ ] 安装 PaddlePaddle + PaddleOCR 环境
- [ ] 运行 `first_run.py`，对4类PDF完成冒烟测试
- [ ] 人工对比识别结果，记录漏识别 / 错识别字段

### Week 3-4（字段规则定义）
- [ ] 为4类报告各建立 JSON 模板配置文件
- [ ] 编写表格提取脚本，验证化学成分表 / 力学性能表的行列解析
- [ ] 确定各字段置信度阈值

### Week 5-7（标注 + 训练）
- [ ] 安装 PPOCRLabel，完成铸件报告标注（≥80页）
- [ ] 完成锻件报告标注（≥60页）
- [ ] 训练 KIE 模型，铸件/锻件字段准确率达到 ≥ 95%
- [ ] 导出推理模型

### Week 8（印章检测 + 优化）
- [ ] 集成印章检测，验证遮挡标记功能
- [ ] 调优置信度阈值，平衡自动通过率与复核率
- [ ] 完成化学成分 / 机械性能复检报告的识别集成

### Week 9-10（测试 + 验收）
- [ ] 用真实报告批量测试（≥50份/类型）
- [ ] 数字字段准确率达到 ≥ 98%
- [ ] 压力测试：200份/天处理能力验证
- [ ] 输出格式验收（JSON + Excel）

---

*本文档持续更新，如有变更请同步修改版本号和日期。*
