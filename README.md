# 农作物的种植策略（2024 数模竞赛 C 题）

本项目用于求解 2024 年高教社杯全国大学生数学建模竞赛 C 题：农作物的种植策略。项目基于题目提供的 `附件1.xlsx`、`附件2.xlsx` 和 `附件3` 中的结果模板，建立农作物种植优化模型，并输出 2024—2030 年各地块、各季节、各作物的种植方案。

本项目不是深度学习项目，也不是简单的随机模拟。核心方法是数学规划、鲁棒优化和多情景风险评估。代码主要完成三件事：读取并清洗附件数据，建立优化模型，最后将结果写回 Excel 模板。

## 一、项目任务说明

题目要求根据乡村现有耕地、农作物、2023 年种植情况、作物亩产量、成本、销售价格等数据，制定 2024—2030 年的农作物种植方案。优化目标是在满足种植制度、作物适种性、轮作要求、豆类种植要求等约束的前提下，提高总体收益，并控制滞销和风险。

本项目对应解决三个问题：

1. 问题一：在未来参数稳定不变的条件下，分别考虑两种销售情形，求最优种植方案。
2. 问题二：考虑销售量、亩产量、成本、价格等不确定性，建立鲁棒种植方案。
3. 问题三：进一步考虑作物之间的可替代性、互补性，以及销售量、价格、成本之间的相关性，并与问题二结果进行比较。

## 二、数据来源与基本处理

项目使用题目官方附件：

- `附件1.xlsx`：包含乡村现有耕地信息和农作物基本信息。
- `附件2.xlsx`：包含 2023 年农作物种植情况和 2023 年统计参数。
- `附件3/`：包含结果填写模板。
- `C题.pdf`：题目说明文档。

代码会自动读取附件并完成以下处理：

1. 读取地块名称、地块类型、地块面积。
2. 读取作物编号、作物名称、作物类型。
3. 读取 2023 年种植情况，处理合并单元格造成的空白地块名称。
4. 读取亩产量、种植成本、销售单价，并将价格区间转为均值。
5. 根据 2023 年种植面积和亩产量推算各作物的基础预期销售量。
6. 根据题目说明补全智慧大棚第一季蔬菜参数。

## 三、三个问题的建模与解法

### 1. 问题一：确定性条件下的种植优化

对应脚本：

```bash
solve_problem1_milp_fast.py
```

问题一假设 2024—2030 年各类参数与 2023 年保持一致，包括预期销售量、亩产量、种植成本、销售价格等。

问题一分为两个情形：

- 情形 1：超过预期销售量的部分滞销浪费，不产生收入。
- 情形 2：超过预期销售量的部分按 2023 年售价的 50% 降价出售。

模型采用混合整数线性规划 MILP。主要变量包括：

```text
x[年份, 地块, 季节, 作物]：种植面积
z[年份, 地块, 季节, 作物]：是否种植该作物
sale[年份, 地块, 季节, 作物]：正常价格销售量
```

主要约束包括：

1. 地块面积约束：每个地块每季种植面积不能超过该地块面积。
2. 地块适种约束：不同地块类型只能种植题目允许的作物。
3. 季节制度约束：
   - 平旱地、梯田、山坡地每年种一季粮食作物。
   - 水浇地可以选择单季水稻，或者两季蔬菜。
   - 普通大棚第一季种蔬菜，第二季种食用菌。
   - 智慧大棚两季均可种蔬菜。
4. 不连续重茬约束：同一地块不能连续种植同一种作物。
5. 豆类约束：每个地块从 2023 年开始，每三年内至少种植一次豆类作物。
6. 最小种植面积约束：避免结果出现过小的碎片化面积。
7. 每季最多作物数约束：控制同一地块同一季种植作物过度分散。

目标函数为最大化总利润：

```text
总利润 = 销售收入 - 种植成本
```

输出文件：

```text
output/result1_1.xlsx
output/result1_2.xlsx
output/problem1_summary.xlsx
```

### 2. 问题二：考虑不确定性和风险的鲁棒优化

对应脚本：

```bash
solve_problem2_robust_milp.py
```

问题二在问题一基础上考虑未来参数变化和风险。代码通过多情景模拟生成 2024—2030 年参数，再用鲁棒收益进行优化。

题目给出的不确定性规则包括：

1. 小麦、玉米预期销售量每年增长 5%—10%。
2. 其他作物预期销售量相对 2023 年上下波动 ±5%。
3. 作物亩产量每年上下波动 ±10%。
4. 种植成本平均每年增长约 5%。
5. 粮食类作物销售价格基本稳定。
6. 蔬菜类作物销售价格平均每年增长约 5%。
7. 食用菌销售价格每年下降 1%—5%。
8. 羊肚菌销售价格每年下降 5%。

模型做法：

1. 生成多个随机情景。
2. 在每个情景下计算作物的销售量、亩产量、价格和成本。
3. 使用低分位数收益作为鲁棒收益参数，避免方案只在乐观情景下表现好。
4. 使用 `scipy.optimize.milp` 建立并求解整数规划模型。
5. 输出问题二的种植方案和情景摘要。

输出文件：

```text
output/result2.xlsx
output/problem2_summary.xlsx
output/problem2_solution_long.csv
```

### 3. 问题三：考虑相关性、替代性、互补性的扩展模型

对应脚本：

```bash
solve_problem3_correlated_robust_fast.py
```

问题三在问题二基础上进一步考虑作物之间的关系，不再把每种作物完全独立处理。

模型主要加入三类因素：

1. 相关性：销售量、价格、成本、亩产量之间不完全独立，而是通过相关情景模拟共同变化。
2. 替代性：部分作物之间存在市场替代关系，同一替代组共享市场容量约束，避免同类作物同时过量种植。
3. 互补性：豆类作物具有轮作改良作用，对后续非豆类作物设置一定协同收益。

模型还会读取问题二输出的：

```text
output/problem2_solution_long.csv
```

如果该文件存在，第三问会自动与问题二方案进行多情景风险比较；如果不存在，则跳过比较分析，但仍会输出第三问方案。

输出文件：

```text
output/result3.xlsx
output/problem3_solution_long.csv
output/problem3_sold.csv
output/problem3_solver_status.csv
output/problem3_summary.xlsx
output/problem3_evaluation_by_scenario.csv
output/problem3_risk_summary.csv
```

## 四、项目目录说明

推荐目录结构如下：

```text
农作物的种植策略（2024数模竞赛C题）/
├─ .vscode/                              # VS Code 配置，可选
├─ 附件3/                                # 题目提供的结果模板
│  ├─ result1_1.xlsx
│  ├─ result1_2.xlsx
│  └─ result2.xlsx
├─ output/                               # 程序输出目录，运行后自动生成或更新
├─ .gitignore                            # Git 忽略规则
├─ README.md                             # 项目说明文档
├─ requirements.txt                      # pip 第三方库依赖
├─ 附件1.xlsx                            # 题目附件：地块和作物信息
├─ 附件2.xlsx                            # 题目附件：2023 种植情况和统计数据
├─ C题.pdf                               # 题目 PDF
├─ solve_problem1_milp_fast.py           # 问题一：确定性 MILP 快速版
├─ solve_problem2_robust_milp.py         # 问题二：多情景鲁棒 MILP
└─ solve_problem3_correlated_robust_fast.py # 问题三：相关性、替代性、互补性模型
```

各文件作用如下：

| 文件或目录 | 作用 |
|---|---|
| `附件1.xlsx` | 官方数据，包含地块信息和作物信息 |
| `附件2.xlsx` | 官方数据，包含 2023 年种植情况和统计参数 |
| `附件3/` | 官方结果模板，代码会将结果写入模板结构 |
| `solve_problem1_milp_fast.py` | 求解问题一的两个情形，并输出 `result1_1.xlsx`、`result1_2.xlsx` |
| `solve_problem2_robust_milp.py` | 求解问题二，输出鲁棒种植方案 |
| `solve_problem3_correlated_robust_fast.py` | 求解问题三，并与问题二方案进行风险比较 |
| `output/` | 存放所有运行结果 |
| `requirements.txt` | pip 安装依赖用的库列表 |
| `README.md` | 项目说明和运行教程 |

## 五、Python 版本与第三方库

推荐 Python 版本：

```text
Python 3.11
```

主要第三方库：

| 库 | 作用 |
|---|---|
| `pandas` | 读取和整理 Excel 表格数据 |
| `openpyxl` | 写回 Excel 模板 |
| `numpy` | 数值计算和随机情景生成 |
| `scipy` | 使用 `scipy.optimize.milp` 求解问题二、问题三 |
| `pulp` | 建立和求解问题一 MILP 模型 |
| `highspy` | HiGHS 求解器接口，加速 MILP 求解 |

## 六、拉取代码与运行教程

### 1. 拉取项目代码

cmd里运行

```bash
git clone https://gitee.com/yu--xin--lei/crop-planting-optimization-2024-c
cd crop-planting-optimization-2024-c
```

如果仓库中没有包含官方附件，需要手动把以下文件放到项目根目录：

```text
附件1.xlsx
附件2.xlsx
C题.pdf
附件3/result1_1.xlsx
附件3/result1_2.xlsx
附件3/result2.xlsx
```

放置完成后，目录结构应与本文第四部分一致。

### 2. 使用 Conda 创建环境

推荐使用 Conda，尤其是在 Windows 上更稳定。

```bash
conda create -n crop2024 python=3.11 -y
conda activate crop2024
conda install -c conda-forge pandas openpyxl numpy scipy pulp highspy coin-or-cbc -y
```

安装完成后可以验证：

```bash
python -c "import pandas, openpyxl, numpy, scipy, pulp, highspy; print('环境正常')"
```

### 3. 没有 Conda 时的解决方式

如果没有 Conda，可以用 Python 自带的虚拟环境 `venv`。

Windows PowerShell：

```bash
python -m venv .venv
.\.venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
```

macOS / Linux：

```bash
python3.11 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
pip install -r requirements.txt
```

安装完成后验证：

```bash
python -c "import pandas, openpyxl, numpy, scipy, pulp, highspy; print('环境正常')"
```

说明：

- Conda 用户建议安装 `coin-or-cbc`，作为 PuLP 的备用求解器。
- pip 用户通常无法直接通过 `requirements.txt` 安装 `coin-or-cbc`，因此主要依赖 `highspy` 和 PuLP 自带/可调用的求解器。
- 如果问题一 CBC 很慢，优先使用 `--solver auto` 或 `--solver highs`。

## 七、运行顺序

建议按问题顺序运行。

### 1. 运行问题一

快速正式版：

```bash
python solve_problem1_milp_fast.py --solver auto --time-limit 300 --gap 0.02
```

更严格版本：

```bash
python solve_problem1_milp_fast.py --solver auto --time-limit 600 --gap 0.01
```

输出：

```text
output/result1_1.xlsx
output/result1_2.xlsx
output/problem1_summary.xlsx
```

### 2. 运行问题二

默认运行：

```bash
python solve_problem2_robust_milp.py
```

正式建议运行：

```bash
python solve_problem2_robust_milp.py --scenarios 200 --time-limit 600 --gap 0.02
```

输出：

```text
output/result2.xlsx
output/problem2_summary.xlsx
output/problem2_solution_long.csv
```

### 3. 运行问题三

默认运行：

```bash
python solve_problem3_correlated_robust_fast.py
```

正式建议运行：

```bash
python solve_problem3_correlated_robust_fast.py --scenarios 300 --time-limit 900 --gap 0.02
```

输出：

```text
output/result3.xlsx
output/problem3_solution_long.csv
output/problem3_sold.csv
output/problem3_solver_status.csv
output/problem3_summary.xlsx
output/problem3_evaluation_by_scenario.csv
output/problem3_risk_summary.csv
```

## 八、常见问题

### 1. 报错：找不到 `附件1.xlsx` 或 `附件2.xlsx`

原因通常是当前运行目录不对，或者附件没有放在项目根目录。

解决方式：

```bash
cd "农作物的种植策略（2024数模竞赛C题）"
python solve_problem1_milp_fast.py
```

或者指定项目路径：

```bash
python solve_problem1_milp_fast.py --base "D:\你的路径\农作物的种植策略（2024数模竞赛C题）"
```

### 2. 报错：找不到 `附件3/result2.xlsx`

问题二和问题三需要使用 `附件3/result2.xlsx` 作为模板。请检查 `附件3` 文件夹是否存在，并且文件名没有被修改。

### 3. 问题一运行太慢

可以使用较宽松的求解参数：

```bash
python solve_problem1_milp_fast.py --solver auto --time-limit 300 --gap 0.02
```

如果电脑性能较好，可以增加时间并缩小 gap：

```bash
python solve_problem1_milp_fast.py --solver auto --time-limit 900 --gap 0.01
```

数学建模中通常不需要为了证明最后 0.x% 的全局最优性无限等待。只要求解状态可行、gap 合理、约束检查通过，就可以作为正式可交付结果。

### 4. pip 安装失败

先升级 pip：

```bash
python -m pip install --upgrade pip setuptools wheel
pip install -r requirements.txt
```

如果仍然失败，建议改用 Conda：

```bash
conda create -n crop2024 python=3.11 -y
conda activate crop2024
conda install -c conda-forge pandas openpyxl numpy scipy pulp highspy coin-or-cbc -y
```

### 5. 中文路径问题

Python 3 对中文路径支持较好，但命令行中路径最好加引号，例如：

```bash
python solve_problem1_milp_fast.py --base "D:\CodeRepository\农作物的种植策略（2024数模竞赛C题）"
```

## 九、结果说明

程序输出的 Excel 文件位于 `output/` 目录下。主要结果文件如下：

| 文件 | 含义 |
|---|---|
| `result1_1.xlsx` | 问题一情形 1：超过销售量部分滞销浪费 |
| `result1_2.xlsx` | 问题一情形 2：超过销售量部分按 50% 售价出售 |
| `result2.xlsx` | 问题二鲁棒优化种植方案 |
| `result3.xlsx` | 问题三相关性、替代性、互补性下的种植方案 |
| `problem1_summary.xlsx` | 问题一求解摘要和明细 |
| `problem2_summary.xlsx` | 问题二参数、收益和求解摘要 |
| `problem3_risk_summary.csv` | 问题三与问题二的风险比较摘要 |
| `problem3_evaluation_by_scenario.csv` | 多情景下方案表现明细 |

## 十、建模边界说明

本项目尽量采用正式建模方式，而不是 baseline。需要说明的是，问题二和问题三涉及未来不确定性、风险偏好、相关性、替代性和互补性，这些内容在题目中没有给出唯一数值标准，因此模型需要做合理假设。

本项目采用的处理方式是：

1. 对题目给出明确范围的参数，严格按照范围生成情景。
2. 对风险使用鲁棒低分位数收益建模，避免过度乐观。
3. 对替代性使用替代组市场容量约束。
4. 对互补性使用豆类轮作协同收益刻画。
5. 对问题三结果与问题二结果进行多情景对比，而不是只给单一收益值。

因此，本项目结果属于可复现、可解释的正式优化方案；如果需要进一步提高严谨性，可以增加情景数量、延长求解时间、缩小 MIP gap，并在论文中对参数设定做敏感性分析。
