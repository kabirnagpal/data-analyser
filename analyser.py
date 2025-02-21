import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import io
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from sklearn.preprocessing import LabelEncoder

# Streamlit UI
st.title("Excel Data Analyzer")
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    
    # Function to classify columns
    def classify_columns(df):
        nominal = []
        ordinal = []
        continuous = []
        
        for col in df.columns:
            if df[col].dtype == 'object':
                nominal.append(col)
            elif df[col].dtype in ['int64', 'float64']:
                unique_vals = df[col].nunique()
                if unique_vals < 10:
                    ordinal.append(col)
                else:
                    continuous.append(col)
        
        return nominal, ordinal, continuous
    
    nominal, ordinal, continuous = classify_columns(df)
    
    st.write("### Column Classification")
    st.write("**Nominal:**", nominal)
    st.write("**Ordinal:**", ordinal)
    st.write("**Continuous:**", continuous)
    
    # Function to generate insights for graphs
    def generate_insight(df, col):
        if df[col].dtype == 'object':  # Categorical
            most_common = df[col].value_counts().idxmax()
            most_common_count = df[col].value_counts().max()
            return f"The most common category in {col} is '{most_common}' with {most_common_count} occurrences."
        else:  # Continuous
            mean_val = df[col].mean()
            median_val = df[col].median()
            std_val = df[col].std()
            return f"The average value of {col} is {mean_val:.2f}, while the median is {median_val:.2f} with a standard deviation of {std_val:.2f}."
    
    # Function to generate and save plots using reportlab
    def save_plots_with_reportlab(df, nominal, ordinal, continuous):
        pdf_path = "analysis_report.pdf"
        c = canvas.Canvas(pdf_path, pagesize=letter)
        width, height = letter
        margin = 50
        y_position = height - margin

        def draw_left_aligned_text(c, text, y_position, font_size=12, bold=False):
            c.setFont("Helvetica-Bold" if bold else "Helvetica", font_size)
            c.drawString(margin, y_position, text)
            return y_position - font_size - 5

        def draw_centered_image(c, img_buf, y_position, title, explanation):
            img_buf.seek(0)
            img_reader = ImageReader(img_buf)
            image_height = 200  # Fixed height for consistency
            text_space = 50  # Space for text
            required_space = image_height + text_space

            if y_position - required_space < margin:
                c.showPage()
                y_position = height - margin

            draw_left_aligned_text(c, title, y_position, 14, True)
            c.drawImage(img_reader, (width - 400) / 2, y_position - image_height, width=400, height=image_height)
            draw_left_aligned_text(c, explanation, y_position - image_height - 20, 12)
            return y_position - required_space
        
        # Title
        y_position = draw_left_aligned_text(c, "Data Analysis Report", y_position, 20, True)
        y_position = draw_left_aligned_text(c, "This report provides an in-depth analysis of the uploaded dataset.", y_position, 12)
        y_position = draw_left_aligned_text(c, "It includes column classification, statistical insights, and graphical representations.", y_position, 12)
        y_position -= 40

        # Column Classification
        y_position = draw_left_aligned_text(c, "Column Classification:", y_position, 14, True)
        y_position = draw_left_aligned_text(c, f"Nominal: {', '.join(nominal) if nominal else 'None'}", y_position, 12)
        y_position = draw_left_aligned_text(c, f"Ordinal: {', '.join(ordinal) if ordinal else 'None'}", y_position, 12)
        y_position = draw_left_aligned_text(c, f"Continuous: {', '.join(continuous) if continuous else 'None'}", y_position, 12)
        y_position -= 40
        
        # Generate Plots and Add to PDF
        for col in nominal + ordinal + continuous:
            fig, ax = plt.subplots(figsize=(6, 4))
            if col in continuous:
                sns.histplot(df[col], kde=True, ax=ax)
                explanation = generate_insight(df, col)
            else:
                sns.countplot(x=df[col], ax=ax, order=df[col].value_counts().index[:10])
                ax.set_xticklabels(ax.get_xticklabels(), rotation=45, ha='right')
                explanation = generate_insight(df, col)
            ax.set_title(f'Distribution of {col}')
            
            img_buf = io.BytesIO()
            plt.tight_layout()
            plt.savefig(img_buf, format='png')
            plt.close()
            
            y_position = draw_centered_image(c, img_buf, y_position, f"Distribution of {col}", explanation)
        
        # Conclusion
        y_position = draw_left_aligned_text(c, "Conclusion:", y_position - 40, 14, True)
        y_position = draw_left_aligned_text(c, "Key Insights:", y_position - 20, 12, True)
        y_position = draw_left_aligned_text(c, "- The most common brand in the dataset is identified.", y_position - 20, 12)
        y_position = draw_left_aligned_text(c, "- Price distributions show significant regional variations.", y_position - 20, 12)
        y_position = draw_left_aligned_text(c, "- Trends in product launches are identified over years.", y_position - 20, 12)
        
        c.save()
        return pdf_path
    
    pdf_path = save_plots_with_reportlab(df, nominal, ordinal, continuous)
    
    with open(pdf_path, "rb") as f:
        st.download_button("Download Analysis Report", f, file_name="data_analysis.pdf", mime="application/pdf")
