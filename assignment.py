import pandas as pd
import matplotlib.pyplot as plt
import plotly.graph_objects as go
import networkx as nx
from math import cos, sin, pi
from docx import Document
from docx.shared import Inches
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Image, Spacer, Paragraph
from reportlab.lib.styles import getSampleStyleSheet

# ----------------------
# 1. Bar Graph
# ----------------------

def create_bar_chart():
    df = pd.read_csv('bar_assignment.csv')
    df.replace({0: 'No', 1: 'Yes'}, inplace=True)
    df_grouped = df.groupby(['LABEL', 'COUNT']).size().unstack(fill_value=0)

    fig, ax = plt.subplots(figsize=(8, 5)) 
    bar_plot = df_grouped.plot(
        kind='barh', 
        stacked=True, 
        ax=ax, 
        color=['red', 'blue']
    )

    ax.set_title('Bar Graph', loc='left', fontsize=12, pad=10)
    ax.set_xlabel('Count', fontsize=10)
    ax.set_ylabel('Label', fontsize=10)

    ax.legend(
        title='Legend', 
        loc='upper left',
        bbox_to_anchor=(1, 1),
        frameon=True,
        fontsize=9
    )

    for container in ax.containers:
        labels = [f'{int(value)}' if value > 0 else '' for value in container.datavalues]
        ax.bar_label(container, labels=labels, label_type='center', fontsize=8, color='white')

    plt.tight_layout(rect=[0, 0, 1, 0.95])
    plt.savefig('bar_chart.png', dpi=300, bbox_inches='tight')
    plt.close()

# ----------------------
# 2. Sankey Diagram
# ----------------------

def create_sankey():
    df = pd.read_csv('sankey_assignment.csv')
    
    left_nodes = ['OMP', 'PS', 'CNP', 'NCDM', 'RGS', 'NRP', 'PEC', 'NMCCC']
    middle_nodes = ['I', 'S', 'D', 'N', 'F']
    right_nodes = ['Aca', 'Oth', 'Reg']
    
    nodes = left_nodes + middle_nodes + right_nodes
    node_indices = {node: i for i, node in enumerate(nodes)}

    source = []
    target = []
    value = []
    

    for _, row in df.iterrows():
        label = row['LABEL'] 
        for col in left_nodes:
            if row[col] > 0:
                source.append(node_indices[col])  
                target.append(node_indices[label])  
                value.append(row[col])
    
    for _, row in df.iterrows():
        label = row['LABEL'] 
        for col in right_nodes:
            if row[col] > 0:
                source.append(node_indices[label])
                target.append(node_indices[col])
                value.append(row[col])
    
    fig = go.Figure(go.Sankey(
        node=dict(
            label=nodes,
            color=['blue'] * len(left_nodes) + ['green'] * len(middle_nodes) + ['orange'] * len(right_nodes)
        ),
        link=dict(source=source, target=target, value=value)
    ))
    fig.update_layout(title_text="Sankey Diagram", font_size=10)
    fig.write_image("sankey.png", scale=2)

# ----------------------
# 3. Network Graph
# ----------------------
def create_network():
    df = pd.read_csv('networks_assignment.csv')
    G = nx.Graph()

    for _, row in df.iterrows():
        source = row['LABELS']
        for target, weight in row[1:].items():
            if weight > 0:
                G.add_edge(source, target, weight=weight)
    
    pentagon = ['D', 'F', 'I', 'N', 'S']
    green = ['BIH', 'GEO', 'ISR', 'MNE', 'SRB', 'CHE', 'TUR', 'UKR', 'GBR', 'AUS', 'HKG', 'USA']
    yellow = ['AUT', 'BEL', 'BGR', 'HRV', 'CZE', 'EST', 'FRA', 'DEU', 'GRC', 'HUN', 'IRL', 'ITA', 'LVA', 'LUX', 'NLD', 'PRT', 'ROU', 'SVK', 'SVN', 'ESP']
    
    node_colors = []
    for node in G.nodes:
        if node in pentagon:
            node_colors.append('blue')
        elif node in green:
            node_colors.append('green')
        elif node in yellow:
            node_colors.append('yellow')
        else:
            node_colors.append('red')
    
    pos = {}
    angles = [pi/2 + 2*pi*i/5 for i in range(5)]
    for i, node in enumerate(pentagon):
        pos[node] = (cos(angles[i]), sin(angles[i]))
    other_nodes = [n for n in G.nodes if n not in pentagon]
    other_pos = nx.circular_layout(other_nodes, scale=2)
    pos.update(other_pos)
    
    plt.figure(figsize=(10, 10))
    nx.draw(G, pos, with_labels=True, node_color=node_colors, edge_color='gray')
    

    plt.text(
        0.5, 1.05, 'Network Graph Label', 
        transform=plt.gca().transAxes, 
        fontsize=12, 
        ha='center', 
        va='center'
    )
    
    plt.savefig('network.png', dpi=300, bbox_inches='tight')
    plt.close()

# ----------------------
# 4. Collate Graphs
# ----------------------
def collate_graphs():
    fig = plt.figure(figsize=(15, 10))
    
    ax1 = plt.subplot2grid((2, 2), (0, 0))  # Bar Chart
    ax2 = plt.subplot2grid((2, 2), (1, 0))  # Sankey Diagram
    ax3 = plt.subplot2grid((2, 2), (0, 1), rowspan=2)  # Network Graph
    
    bar_img = plt.imread('bar_chart.png')
    sankey_img = plt.imread('sankey.png')
    network_img = plt.imread('network.png')
    
    ax1.imshow(bar_img)
    ax1.axis('off')
    
    ax2.imshow(sankey_img)
    ax2.axis('off')
    
    ax3.imshow(network_img)
    ax3.axis('off')
    
    plt.tight_layout()
    plt.savefig('collated.png', dpi=300, bbox_inches='tight')
    plt.close()

# ----------------------
# 5. Create PDF Sample
# ----------------------
def create_pdf():
    pdf_file = "collated_graphs.pdf"
    doc = SimpleDocTemplate(pdf_file, pagesize=letter)
    

    styles = getSampleStyleSheet()
    label_text = Paragraph("Collated Graphs", styles["Title"])
    collated_img = Image("collated.png", width=400, height=250)
    
    content = [
        label_text,
        Spacer(1, 10),
        collated_img
    ]
    
    doc.build(content)

# ----------------------
# Run All Functions
# ----------------------
if __name__ == "__main__":
    create_bar_chart()
    create_sankey()
    create_network()
    collate_graphs()
    create_pdf()