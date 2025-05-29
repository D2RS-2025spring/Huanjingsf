import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.colors import LinearSegmentedColormap
from matplotlib.legend import Legend

# ==================== 参数配置 ====================
CUSTOM_COLORS = [
    '#FCFEA6', '#FFE98A', '#FEBA6C', '#FB8758', 
    '#EF5B49', '#D73C47', '#AD243E', '#7F163C', 
    '#440D2C'
]
FIGURE_SIZE = (20, 12)  # 优化宽高比
POINT_SIZE_RANGE = (20, 180)
COLOR_RANGE = (0, 2.5)
FONT_CONFIG = {
    'family': 'Arial',
    'size': 11,
    'color': '#2E2E2E'
}

# ==================== 数据加载 ====================
def load_data():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(
        current_dir,
        "CRC_scRNAseq-1.0.0",
        "41467_2022_29366_MOESM9_ESM.xlsx"
    )

    if not os.path.exists(file_path):
        raise FileNotFoundError(f"文件路径不存在: {file_path}")

    try:
        df = pd.read_excel(file_path, sheet_name='Figure 1c', engine='openpyxl')
        # 列名标准化处理
        df = df.rename(columns={
            'features.plot': 'Gene',
            'id': 'CellType',
            'pct.exp': 'pct.exp',
            'avg.exp.scaled': 'avg.exp.scaled'
        })
        # 数据验证
        required_columns = ['Gene', 'CellType', 'pct.exp', 'avg.exp.scaled']
        if not all(col in df.columns for col in required_columns):
            missing = [col for col in required_columns if col not in df.columns]
            raise KeyError(f"缺失必要列: {missing}")
            
        df = df[required_columns]
        df['pct.exp'] = pd.to_numeric(df['pct.exp'], errors='coerce')
        df['avg.exp.scaled'] = pd.to_numeric(df['avg.exp.scaled'], errors='coerce')
        df = df.dropna()
        return df
    except Exception as e:
        print(f"数据加载失败: {str(e)}")
        exit()

# ==================== 可视化引擎 ====================
def create_scatterplot(df):
    fig = plt.figure(figsize=FIGURE_SIZE)
    gs = plt.GridSpec(2, 1, height_ratios=[8, 1], hspace=0.1)
    main_ax = fig.add_subplot(gs[0])
    legend_ax = fig.add_subplot(gs[1])
    legend_ax.axis('off')

    # 主图绘制
    cmap = LinearSegmentedColormap.from_list('nature', CUSTOM_COLORS, N=512)
    scatter = sns.scatterplot(
        data=df,
        x="Gene",
        y="CellType",
        size="pct.exp",
        hue="avg.exp.scaled",
        palette=cmap,
        sizes=POINT_SIZE_RANGE,
        hue_norm=COLOR_RANGE,
        size_norm=(0, 100),
        alpha=0.95,
        linewidth=0.5,
        edgecolor=FONT_CONFIG['color'],
        legend=False,
        ax=main_ax
    )

    # 主图样式优化
    main_ax.spines['top'].set_visible(False)
    main_ax.spines['right'].set_visible(False)
    for spine in ['bottom', 'left']:
        main_ax.spines[spine].set_color('#808080')
    main_ax.tick_params(axis='both', colors='#808080')
    
    plt.setp(main_ax.get_xticklabels(), rotation=90, 
            fontsize=FONT_CONFIG['size'], 
            fontname=FONT_CONFIG['family'],
            color=FONT_CONFIG['color'])
    plt.setp(main_ax.get_yticklabels(), 
            fontsize=FONT_CONFIG['size'],
            fontname=FONT_CONFIG['family'],
            color=FONT_CONFIG['color'])
    
    main_ax.set_xlabel("Gene Markers", 
                     fontsize=FONT_CONFIG['size']+2,
                     labelpad=12,
                     fontname=FONT_CONFIG['family'],
                     color=FONT_CONFIG['color'])
    main_ax.set_ylabel("Cell Types", 
                      fontsize=FONT_CONFIG['size']+2,
                      labelpad=12,
                      fontname=FONT_CONFIG['family'],
                      color=FONT_CONFIG['color'])
    main_ax.grid(axis='y', linestyle='--', linewidth=0.5, alpha=0.4, color='#D3D3D3')

    # ===== 专业级组合图例系统 =====
    # 颜色条配置
    cax = fig.add_axes([0.3, 0.09, 0.4, 0.03])
    cbar = plt.colorbar(scatter.collections[0], cax=cax, orientation='horizontal')
    cbar.set_ticks([0, 1.25, 2.5])
    cbar.ax.tick_params(
        labelsize=FONT_CONFIG['size']-1,
        length=4,
        width=1,
        color=FONT_CONFIG['color'],
        pad=2
    )
    cbar.outline.set_edgecolor(FONT_CONFIG['color'])
    cbar.ax.set_xticklabels(
        ['Low', 'Medium', 'High'],
        fontname=FONT_CONFIG['family'],
        color=FONT_CONFIG['color']
    )
    cbar.ax.set_title(
        'Expression Level',
        fontsize=FONT_CONFIG['size'],
        fontname=FONT_CONFIG['family'],
        color=FONT_CONFIG['color'],
        pad=8,
        loc='left'
    )

    # 大小图例系统
    legend_sizes = [10, 30, 50, 70, 90]
    size_scaler = lambda x: np.sqrt(x/100 * (POINT_SIZE_RANGE[1]-POINT_SIZE_RANGE[0])) + POINT_SIZE_RANGE[0]/2
    
    handles = [
        plt.Line2D([0], [0],
                  marker='o',
                  markersize=size_scaler(s)/15,  # 比例缩放
                  linewidth=0.5,
                  markerfacecolor='#808080',
                  markeredgecolor=FONT_CONFIG['color']) 
        for s in legend_sizes
    ]
    
    # 创建专业图例
    legend = Legend(
        legend_ax,
        handles,
        [f"{s}%" for s in legend_sizes],
        title='% Expressed',
        frameon=False,
        title_fontsize=FONT_CONFIG['size'],
        fontsize=FONT_CONFIG['size']-1,
        handletextpad=1.2,
        columnspacing=2.5,
        ncol=len(legend_sizes),
        loc='center',
        bbox_to_anchor=(0.5, 0.5)
    )
    legend_ax.add_artist(legend)
    
    # 样式统一设置
    for text in legend.get_texts():
        text.set_color(FONT_CONFIG['color'])
        text.set_fontfamily(FONT_CONFIG['family'])
    legend.get_title().set_color(FONT_CONFIG['color'])
    legend.get_title().set_fontfamily(FONT_CONFIG['family'])

    # 布局微调
    plt.subplots_adjust(left=0.08, right=0.92, bottom=0.15, top=0.95)
    plt.show()

# ==================== 主程序 ====================
if __name__ == "__main__":
    data_df = load_data()
    create_scatterplot(data_df)
