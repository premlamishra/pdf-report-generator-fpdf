import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF
from datetime import datetime

#1: Loading and cleaning the dataset
df = pd.read_excel("superstore_sample.xlsx")
df.columns = df.columns.str.strip().str.replace(" ", "_").str.lower()

# Summary by category
summary_df = df.groupby('product_category')[['unit_price', 'shipping_cost']].sum().reset_index()

#2: Generating Charts and Save as PNG

# Chart 1: Bar chart – Total Sales by Product Category
plt.figure(figsize=(6,4))
summary_df.plot(x='product_category', y='unit_price', kind='bar', legend=False, color='#69A1F4')
plt.title("Total Sales by Product Category")
plt.ylabel("Total Sales (Unit Price in $)")
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig("category_chart.png")
plt.close()

# Chart 2: Pie chart – Shipping Cost by Product Category
plt.figure(figsize=(5,5))
df.groupby('product_category')['shipping_cost'].sum().plot(
    kind='pie', autopct='%1.1f%%', startangle=140, colors=plt.cm.Pastel1.colors)
plt.title("Shipping Cost Distribution by Category")
plt.ylabel('')
plt.tight_layout()
plt.savefig("profit_pie.png")
plt.close()

# Chart 3: Line chart – Sales Trend
plt.figure(figsize=(6,4))
df['unit_price'].plot(kind='line', color='#7FB77E', marker='o')
plt.title("Sales Trend Across Orders")
plt.xlabel("Entry Index")
plt.ylabel("Unit Price in $")
plt.tight_layout()
plt.savefig("revenue_chart.png")
plt.close()

#3: Creating PDF Report

class PDF(FPDF):
    def header(self):
        self.set_font("Arial", 'B', 16)
        try:
            self.image("logo.png", 10, 8, 20)
        except:
            pass
        self.set_text_color(0, 102, 204)
        self.cell(0, 10, "Superstore Sales Report", ln=True, align='C')
        self.set_font("Arial", '', 10)
        self.set_text_color(0, 0, 0)
        self.cell(0, 10, f"Generated on {datetime.now().strftime('%B %d, %Y')}", ln=True, align='C')
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", 'I', 8)
        self.cell(0, 10, f"Page {self.page_no()} - Confidential", 0, 0, 'C')

    def add_title_page(self):
        self.add_page()
        self.set_font("Arial", 'B', 24)
        self.set_text_color(30, 30, 30)
        self.ln(50)
        self.cell(0, 10, "Superstore PDF Report", ln=True, align='C')
        self.set_font("Arial", '', 14)
        self.set_text_color(100, 100, 100)
        self.cell(0, 10, "Automated Sales and Shipping Summary", ln=True, align='C')
        self.ln(20)
        self.set_font("Arial", 'I', 12)
        self.cell(0, 10, "Prepared by: Premla Mishra", ln=True, align='C')
        self.cell(0, 10, "Intern - BI/Analytics Team", ln=True, align='C')

    def add_summary_table(self, df):
        self.add_page()
        self.set_font("Arial", 'B', 16)
        self.set_text_color(30, 30, 30)
        self.cell(0, 10, "Category-wise Sales & Shipping Summary", ln=True)
        self.set_fill_color(220, 220, 220)
        self.set_font("Arial", 'B', 12)
        self.cell(70, 10, "Category", 1, 0, 'C', True)
        self.cell(50, 10, "Total Sales", 1, 0, 'C', True)
        self.cell(50, 10, "Shipping Cost", 1, 1, 'C', True)
        self.set_font("Arial", '', 12)
        for i, row in df.iterrows():
            fill = i % 2 == 0
            self.set_fill_color(245, 245, 245) if fill else self.set_fill_color(255, 255, 255)
            self.cell(70, 10, str(row['product_category']), 1, 0, 'L', True)
            self.cell(50, 10, f"${row['unit_price']:.2f}", 1, 0, 'R', True)
            self.cell(50, 10, f"${row['shipping_cost']:.2f}", 1, 1, 'R', True)

    def add_chart(self, title, path):
        self.add_page()
        self.set_font("Arial", 'B', 16)
        self.set_text_color(0, 102, 204)
        self.cell(0, 10, title, ln=True)
        self.ln(5)
        try:
            self.image(path, x=(210 - 150) // 2, w=150)
        except:
            self.set_text_color(255, 0, 0)
            self.set_font("Arial", '', 12)
            self.cell(0, 10, f"Failed to load image: {path}", ln=True)

# Step 4: Building the PDF
pdf = PDF()
pdf.set_auto_page_break(auto=True, margin=15)
pdf.add_title_page()
pdf.add_summary_table(summary_df)
pdf.add_chart("Total Sales by Product Category", "category_chart.png")
pdf.add_chart("Shipping Cost Distribution", "profit_pie.png")
pdf.add_chart("Sales Trend Across Orders", "revenue_chart.png")
pdf.output("report.pdf")

print("PDF report generated: report.pdf")
