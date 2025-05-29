import pandas as pd
import matplotlib.pyplot as plt
import os


# 动态获取当前脚本目录
current_dir = os.path.dirname(os.path.abspath(__file__))

# 构建相对路径
file_path = os.path.join(
    current_dir, 
    "CRC_scRNAseq-1.0.0", 
    "41467_2022_29366_MOESM9_ESM.xlsx"
)

# 验证文件存在性
if not os.path.exists(file_path):
    raise FileNotFoundError(f"文件不存在: {file_path}")

# 加载 Excel 文件（指定工作表）
data = pd.read_excel(file_path, sheet_name='Figure 1e_left_panel')

# 将 'Sample' 列转换为有序分类变量
data['Sample'] = pd.Categorical(
    data['Sample'], 
    categories=['p1', 'p2', 'p3', 'p4', 'p5'], 
    ordered=True
)

# 创建透视表
pivot_data = data.pivot(index='CellTypes', columns='Sample', values='Per')

# 设置颜色、绘图参数
colors = ['#00a6d0', '#00dbb0', '#ffa861', '#b000ff', '#ff0062']  # P1-P5 颜色
fig, ax = plt.subplots(figsize=(6, 4))
pivot_data.plot(kind='barh', stacked=True, color=colors, ax=ax, width=0.6)

# 调整坐标轴样式
ax.xaxis.set_tick_params(width=2)
ax.yaxis.set_tick_params(width=2)
ax.spines['right'].set_visible(False)
ax.spines['bottom'].set_visible(False)
ax.xaxis.tick_top()
ax.xaxis.set_label_position('top')

# 设置边框线宽
for axis in ['top', 'left']:
    ax.spines[axis].set_linewidth(2)

# 添加标题和标签
plt.title('Percentage of Cell Types across Different Patients')
plt.xlabel('Percentage of cells (%)')
plt.ylabel('Cell Types')
plt.legend(title='Patients', bbox_to_anchor=(1.05, 1), loc='upper left')

# 显示图像
plt.tight_layout()
plt.show()