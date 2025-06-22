import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import io
import base64
from datetime import datetime
import os

# Page configuration
st.set_page_config(
    page_title="General Excel Analyzer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #1f77b4;
    }
    .upload-section {
        background-color: #e8f4fd;
        padding: 2rem;
        border-radius: 1rem;
        margin: 1rem 0;
        text-align: center;
    }
    .stButton > button {
        width: 100%;
        margin: 0.5rem 0;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'uploaded_file' not in st.session_state:
    st.session_state.uploaded_file = None
if 'sheets_data' not in st.session_state:
    st.session_state.sheets_data = {}
if 'sheet_names' not in st.session_state:
    st.session_state.sheet_names = []
if 'current_sheet' not in st.session_state:
    st.session_state.current_sheet = None
if 'edited_data' not in st.session_state:
    st.session_state.edited_data = {}

def load_excel_data(uploaded_file):
    """Load all sheets from uploaded Excel file"""
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        sheets_data = {}
        
        for sheet_name in excel_file.sheet_names:
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
            sheets_data[sheet_name] = df
            
        return sheets_data, excel_file.sheet_names
    except Exception as e:
        st.error(f"Error loading Excel file: {str(e)}")
        return None, None

def display_file_upload():
    """Display file upload section"""
    st.markdown('<h2 class="main-header">📊 General Excel Analyzer</h2>', unsafe_allow_html=True)
    
    st.markdown("""
    <div class="upload-section">
        <h3>Upload Your Excel File</h3>
        <p>Upload any Excel file (.xlsx, .xls) to analyze and manipulate its data</p>
    </div>
    """, unsafe_allow_html=True)
    
    uploaded_file = st.file_uploader(
        "Choose an Excel file",
        type=['xlsx', 'xls'],
        help="Upload an Excel file to get started"
    )
    
    if uploaded_file is not None:
        st.session_state.uploaded_file = uploaded_file
        sheets_data, sheet_names = load_excel_data(uploaded_file)
        
        if sheets_data:
            st.session_state.sheets_data = sheets_data
            st.session_state.sheet_names = sheet_names
            st.success(f"✅ Successfully loaded {len(sheet_names)} sheets from {uploaded_file.name}")
            return True
    
    return False

def display_overview():
    """Display overview of uploaded Excel file"""
    st.markdown('<h3>📋 File Overview</h3>', unsafe_allow_html=True)
    
    sheets_data = st.session_state.sheets_data
    sheet_names = st.session_state.sheet_names
    
    # Summary metrics
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Total Sheets", len(sheet_names))
    
    with col2:
        total_rows = sum(len(sheets_data[sheet]) for sheet in sheet_names)
        st.metric("Total Rows", total_rows)
    
    with col3:
        total_columns = sum(len(sheets_data[sheet].columns) for sheet in sheet_names)
        st.metric("Total Columns", total_columns)
    
    with col4:
        file_size = len(st.session_state.uploaded_file.getvalue()) / 1024  # KB
        st.metric("File Size", f"{file_size:.1f} KB")
    
    # Sheet information table
    st.subheader("📊 Sheet Information")
    
    sheet_info = []
    for sheet_name in sheet_names:
        df = sheets_data[sheet_name]
        
        # Get data types summary
        dtype_counts = df.dtypes.value_counts()
        dtype_summary = ", ".join([f"{dtype}: {count}" for dtype, count in dtype_counts.items()])
        
        # Get missing values
        missing_values = df.isnull().sum().sum()
        
        sheet_info.append({
            'Sheet Name': sheet_name,
            'Rows': len(df),
            'Columns': len(df.columns),
            'Missing Values': missing_values,
            'Data Types': dtype_summary[:50] + "..." if len(dtype_summary) > 50 else dtype_summary
        })
    
    sheet_df = pd.DataFrame(sheet_info)
    st.dataframe(sheet_df, use_container_width=True)

def display_data_explorer():
    """Display data explorer with sheet selection and data viewing"""
    st.markdown('<h3>📄 Data Explorer</h3>', unsafe_allow_html=True)
    
    sheets_data = st.session_state.sheets_data
    sheet_names = st.session_state.sheet_names
    
    # Sheet selector
    selected_sheet = st.selectbox("Select a sheet to view:", sheet_names)
    
    if selected_sheet:
        df = sheets_data[selected_sheet]
        st.session_state.current_sheet = selected_sheet
        
        # Display basic info
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Rows", len(df))
        with col2:
            st.metric("Columns", len(df.columns))
        with col3:
            missing_values = df.isnull().sum().sum()
            st.metric("Missing Values", missing_values)
        with col4:
            memory_usage = df.memory_usage(deep=True).sum() / 1024  # KB
            st.metric("Memory Usage", f"{memory_usage:.1f} KB")
        
        # Data preview
        st.subheader("Data Preview")
        
        # Show first few rows
        preview_rows = st.slider("Number of rows to preview:", 5, 50, 10)
        st.dataframe(df.head(preview_rows), use_container_width=True)
        
        # Column information
        st.subheader("Column Information")
        col_info = []
        for col in df.columns:
            col_info.append({
                'Column': col,
                'Data Type': str(df[col].dtype),
                'Non-Null Count': df[col].count(),
                'Null Count': df[col].isnull().sum(),
                'Unique Values': df[col].nunique()
            })
        
        col_df = pd.DataFrame(col_info)
        st.dataframe(col_df, use_container_width=True)
        
        # Download options
        st.subheader("Download Options")
        col1, col2 = st.columns(2)
        
        with col1:
            csv = df.to_csv(index=False)
            st.download_button(
                label="Download as CSV",
                data=csv,
                file_name=f"{selected_sheet}.csv",
                mime="text/csv"
            )
        
        with col2:
            # Create Excel buffer
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name=selected_sheet, index=False)
            excel_data = output.getvalue()
            
            st.download_button(
                label="Download as Excel",
                data=excel_data,
                file_name=f"{selected_sheet}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

def display_data_manipulation():
    """Display data manipulation interface"""
    st.markdown('<h3>✏️ Data Manipulation</h3>', unsafe_allow_html=True)
    
    if not st.session_state.current_sheet:
        st.warning("Please select a sheet in the Data Explorer first.")
        return
    
    sheets_data = st.session_state.sheets_data
    current_sheet = st.session_state.current_sheet
    df = sheets_data[current_sheet].copy()
    
    # Store original data
    if current_sheet not in st.session_state.edited_data:
        st.session_state.edited_data[current_sheet] = df.copy()
    
    edited_df = st.session_state.edited_data[current_sheet]
    
    st.subheader(f"Editing: {current_sheet}")
    
    # Manipulation options
    manipulation_type = st.selectbox(
        "Choose manipulation type:",
        ["View/Edit Data", "Filter Data", "Sort Data", "Add/Remove Columns", "Data Cleaning"]
    )
    
    if manipulation_type == "View/Edit Data":
        st.write("**Interactive Data Editor**")
        st.write("Use the data editor below to modify values. Changes will be saved when you click 'Save Changes'.")
        
        # Convert dataframe to ensure all columns are editable
        # Convert all columns to string first to avoid data type restrictions
        editable_df = edited_df.copy()
        for col in editable_df.columns:
            editable_df[col] = editable_df[col].astype(str)
        
        # Create column configuration to ensure all columns are editable
        column_config = {}
        for col in editable_df.columns:
            column_config[col] = st.column_config.TextColumn(
                col,
                help=f"Edit {col}",
                max_chars=None,
                validate=None
            )
        
        # Use st.data_editor with explicit column configuration
        edited_df_result = st.data_editor(
            editable_df,
            num_rows="dynamic",
            use_container_width=True,
            column_config=column_config,
            key=f"editor_{current_sheet}",
            hide_index=True
        )
        
        # Convert back to appropriate data types when saving
        if st.button("Save Changes"):
            # Try to convert back to original data types where possible
            final_df = edited_df_result.copy()
            original_df = edited_df.copy()
            
            for col in final_df.columns:
                if col in original_df.columns:
                    # Try to convert back to original data type
                    try:
                        if original_df[col].dtype in ['int64', 'float64']:
                            # Try to convert to numeric
                            final_df[col] = pd.to_numeric(final_df[col], errors='coerce')
                        elif original_df[col].dtype == 'bool':
                            # Try to convert to boolean
                            final_df[col] = final_df[col].map({'True': True, 'False': False, 'true': True, 'false': False, '1': True, '0': False})
                        else:
                            # Keep as string/object
                            pass
                    except:
                        # If conversion fails, keep as string
                        pass
            
            st.session_state.edited_data[current_sheet] = final_df
            st.success("Changes saved!")
        
        if st.button("Reset to Original"):
            st.session_state.edited_data[current_sheet] = df.copy()
            st.rerun()
    
    elif manipulation_type == "Filter Data":
        st.write("**Filter Data**")
        
        # Column selector for filtering
        filter_column = st.selectbox("Select column to filter:", edited_df.columns)
        
        if filter_column:
            if edited_df[filter_column].dtype in ['object', 'string']:
                # Categorical filtering
                unique_values = edited_df[filter_column].dropna().unique()
                selected_values = st.multiselect(
                    f"Select values to keep in '{filter_column}':",
                    unique_values,
                    default=list(unique_values)
                )
                
                if st.button("Apply Filter"):
                    filtered_df = edited_df[edited_df[filter_column].isin(selected_values)]
                    st.session_state.edited_data[current_sheet] = filtered_df
                    st.success(f"Filtered to {len(filtered_df)} rows!")
            
            else:
                # Numerical filtering
                min_val = float(edited_df[filter_column].min())
                max_val = float(edited_df[filter_column].max())
                
                col1, col2 = st.columns(2)
                with col1:
                    min_filter = st.number_input("Minimum value:", value=min_val, key="min_filter")
                with col2:
                    max_filter = st.number_input("Maximum value:", value=max_val, key="max_filter")
                
                if st.button("Apply Filter"):
                    filtered_df = edited_df[
                        (edited_df[filter_column] >= min_filter) & 
                        (edited_df[filter_column] <= max_filter)
                    ]
                    st.session_state.edited_data[current_sheet] = filtered_df
                    st.success(f"Filtered to {len(filtered_df)} rows!")
    
    elif manipulation_type == "Sort Data":
        st.write("**Sort Data**")
        
        sort_column = st.selectbox("Select column to sort by:", edited_df.columns)
        sort_ascending = st.checkbox("Sort in ascending order", value=True)
        
        if st.button("Apply Sort"):
            sorted_df = edited_df.sort_values(by=sort_column, ascending=sort_ascending)
            st.session_state.edited_data[current_sheet] = sorted_df
            st.success("Data sorted!")
    
    elif manipulation_type == "Add/Remove Columns":
        st.write("**Add/Remove Columns**")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Remove Columns")
            columns_to_remove = st.multiselect(
                "Select columns to remove:",
                edited_df.columns
            )
            
            if st.button("Remove Selected Columns"):
                if columns_to_remove:
                    edited_df = edited_df.drop(columns=columns_to_remove)
                    st.session_state.edited_data[current_sheet] = edited_df
                    st.success(f"Removed {len(columns_to_remove)} columns!")
        
        with col2:
            st.subheader("Add New Column")
            new_column_name = st.text_input("New column name:")
            new_column_value = st.text_input("Default value (leave empty for NaN):")
            
            if st.button("Add Column"):
                if new_column_name:
                    if new_column_value:
                        edited_df[new_column_name] = new_column_value
                    else:
                        edited_df[new_column_name] = np.nan
                    st.session_state.edited_data[current_sheet] = edited_df
                    st.success(f"Added column '{new_column_name}'!")
    
    elif manipulation_type == "Data Cleaning":
        st.write("**Data Cleaning**")
        
        cleaning_options = st.multiselect(
            "Select cleaning operations:",
            ["Remove duplicate rows", "Fill missing values", "Remove rows with missing values", "Convert data types"]
        )
        
        if st.button("Apply Cleaning"):
            cleaned_df = edited_df.copy()
            
            if "Remove duplicate rows" in cleaning_options:
                cleaned_df = cleaned_df.drop_duplicates()
                st.info(f"Removed {len(edited_df) - len(cleaned_df)} duplicate rows")
            
            if "Remove rows with missing values" in cleaning_options:
                original_len = len(cleaned_df)
                cleaned_df = cleaned_df.dropna()
                st.info(f"Removed {original_len - len(cleaned_df)} rows with missing values")
            
            if "Fill missing values" in cleaning_options:
                fill_method = st.selectbox("Fill method:", ["Forward fill", "Backward fill", "Fill with 0", "Fill with mean"])
                
                if fill_method == "Forward fill":
                    cleaned_df = cleaned_df.fillna(method='ffill')
                elif fill_method == "Backward fill":
                    cleaned_df = cleaned_df.fillna(method='bfill')
                elif fill_method == "Fill with 0":
                    cleaned_df = cleaned_df.fillna(0)
                elif fill_method == "Fill with mean":
                    numeric_columns = cleaned_df.select_dtypes(include=[np.number]).columns
                    for col in numeric_columns:
                        cleaned_df[col] = cleaned_df[col].fillna(cleaned_df[col].mean())
            
            st.session_state.edited_data[current_sheet] = cleaned_df
            st.success("Data cleaning completed!")
    
    # Show current data
    st.subheader("Current Data")
    st.dataframe(edited_df, use_container_width=True)
    
    # Save all changes
    st.subheader("Save Changes")
    if st.button("Save All Changes to Session"):
        st.session_state.sheets_data[current_sheet] = edited_df.copy()
        st.success("All changes saved to session!")

def create_visualizations():
    """Create various visualizations for the data"""
    st.markdown('<h3>📈 Data Visualizations</h3>', unsafe_allow_html=True)
    
    sheets_data = st.session_state.sheets_data
    sheet_names = st.session_state.sheet_names
    
    # Sheet selector for visualization
    viz_sheet = st.selectbox("Select sheet for visualization:", sheet_names, key="viz_sheet")
    
    if viz_sheet:
        df = sheets_data[viz_sheet]
        
        # Visualization options
        viz_type = st.selectbox(
            "Choose visualization type:",
            ["Data Distribution", "Correlation Matrix", "Missing Values", "Custom Plot"]
        )
        
        if viz_type == "Data Distribution":
            st.write("**Data Distribution**")
            
            # Select numeric columns
            numeric_columns = df.select_dtypes(include=[np.number]).columns
            
            if len(numeric_columns) > 0:
                selected_column = st.selectbox("Select column:", numeric_columns)
                
                col1, col2 = st.columns(2)
                
                with col1:
                    # Histogram
                    fig_hist = px.histogram(df, x=selected_column, title=f"Distribution of {selected_column}")
                    st.plotly_chart(fig_hist, use_container_width=True)
                
                with col2:
                    # Box plot
                    fig_box = px.box(df, y=selected_column, title=f"Box Plot of {selected_column}")
                    st.plotly_chart(fig_box, use_container_width=True)
            else:
                st.warning("No numeric columns found for distribution analysis.")
        
        elif viz_type == "Correlation Matrix":
            st.write("**Correlation Matrix**")
            
            numeric_df = df.select_dtypes(include=[np.number])
            
            if len(numeric_df.columns) > 1:
                corr_matrix = numeric_df.corr()
                
                fig = px.imshow(
                    corr_matrix,
                    title="Correlation Matrix",
                    color_continuous_scale='RdBu',
                    aspect="auto"
                )
                st.plotly_chart(fig, use_container_width=True)
                
                # Show correlation values
                st.subheader("Correlation Values")
                st.dataframe(corr_matrix, use_container_width=True)
            else:
                st.warning("Need at least 2 numeric columns for correlation analysis.")
        
        elif viz_type == "Missing Values":
            st.write("**Missing Values Analysis**")
            
            # Calculate missing values
            missing_data = df.isnull().sum()
            missing_percent = (missing_data / len(df)) * 100
            
            missing_df = pd.DataFrame({
                'Column': missing_data.index,
                'Missing Count': missing_data.values,
                'Missing Percentage': missing_percent.values
            }).sort_values('Missing Count', ascending=False)
            
            # Bar chart of missing values
            fig = px.bar(
                missing_df,
                x='Column',
                y='Missing Count',
                title="Missing Values by Column"
            )
            st.plotly_chart(fig, use_container_width=True)
            
            # Table of missing values
            st.subheader("Missing Values Summary")
            st.dataframe(missing_df, use_container_width=True)
        
        elif viz_type == "Custom Plot":
            st.write("**Custom Plot**")
            
            plot_type = st.selectbox(
                "Select plot type:",
                ["Scatter Plot", "Bar Chart", "Line Chart", "Pie Chart"]
            )
            
            if plot_type == "Scatter Plot":
                col1, col2 = st.columns(2)
                with col1:
                    x_col = st.selectbox("X-axis:", df.columns)
                with col2:
                    y_col = st.selectbox("Y-axis:", df.columns)
                
                if st.button("Create Scatter Plot"):
                    fig = px.scatter(df, x=x_col, y=y_col, title=f"{y_col} vs {x_col}")
                    st.plotly_chart(fig, use_container_width=True)
            
            elif plot_type == "Bar Chart":
                col1, col2 = st.columns(2)
                with col1:
                    x_col = st.selectbox("X-axis:", df.columns)
                with col2:
                    y_col = st.selectbox("Y-axis (numeric):", df.select_dtypes(include=[np.number]).columns)
                
                if st.button("Create Bar Chart"):
                    fig = px.bar(df, x=x_col, y=y_col, title=f"{y_col} by {x_col}")
                    st.plotly_chart(fig, use_container_width=True)
            
            elif plot_type == "Line Chart":
                col1, col2 = st.columns(2)
                with col1:
                    x_col = st.selectbox("X-axis:", df.columns)
                with col2:
                    y_col = st.selectbox("Y-axis (numeric):", df.select_dtypes(include=[np.number]).columns)
                
                if st.button("Create Line Chart"):
                    fig = px.line(df, x=x_col, y=y_col, title=f"{y_col} over {x_col}")
                    st.plotly_chart(fig, use_container_width=True)
            
            elif plot_type == "Pie Chart":
                col_col = st.selectbox("Select column for pie chart:", df.columns)
                
                if st.button("Create Pie Chart"):
                    value_counts = df[col_col].value_counts()
                    fig = px.pie(values=value_counts.values, names=value_counts.index, title=f"Distribution of {col_col}")
                    st.plotly_chart(fig, use_container_width=True)

def export_data():
    """Export modified data"""
    st.markdown('<h3>💾 Export Data</h3>', unsafe_allow_html=True)
    
    if not st.session_state.edited_data:
        st.info("No modified data to export. Use the Data Manipulation section to make changes.")
        return
    
    st.write("**Export Modified Data**")
    
    # Create Excel file with all modified sheets
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in st.session_state.edited_data.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    excel_data = output.getvalue()
    
    # Generate filename with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"modified_excel_{timestamp}.xlsx"
    
    st.download_button(
        label="Download Modified Excel File",
        data=excel_data,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    # Show summary of changes
    st.subheader("Summary of Changes")
    for sheet_name, df in st.session_state.edited_data.items():
        original_df = st.session_state.sheets_data[sheet_name]
        st.write(f"**{sheet_name}:** {len(original_df)} → {len(df)} rows")

def main():
    # Main title
    # st.markdown('<h1 class="main-header">📊 General Excel Analyzer</h1>', unsafe_allow_html=True)
    
    # File upload section
    if not st.session_state.uploaded_file:
        if display_file_upload():
            st.rerun()
        return
    
    # Sidebar navigation
    st.sidebar.title("Navigation")
    page = st.sidebar.selectbox(
        "Choose a page:",
        ["Overview", "Data Explorer", "Data Manipulation", "Visualizations", "Export Data"]
    )
    
    # Display file info in sidebar
    st.sidebar.markdown("---")
    st.sidebar.markdown(f"**File:** {st.session_state.uploaded_file.name}")
    st.sidebar.markdown(f"**Sheets:** {len(st.session_state.sheet_names)}")
    
    # Upload new file option
    if st.sidebar.button("Upload New File"):
        st.session_state.uploaded_file = None
        st.session_state.sheets_data = {}
        st.session_state.sheet_names = []
        st.session_state.current_sheet = None
        st.session_state.edited_data = {}
        st.rerun()
    
    # Page routing
    if page == "Overview":
        display_overview()
    elif page == "Data Explorer":
        display_data_explorer()
    elif page == "Data Manipulation":
        display_data_manipulation()
    elif page == "Visualizations":
        create_visualizations()
    elif page == "Export Data":
        export_data()

if __name__ == "__main__":
    main() 