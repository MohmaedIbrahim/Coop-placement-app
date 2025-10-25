import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from scipy.cluster.hierarchy import dendrogram, linkage, fcluster
from scipy.spatial.distance import squareform
import pulp
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

# Page config
st.set_page_config(page_title="Co-op Placement Optimizer", layout="wide", page_icon="🎓")

# Initialize session state
if 'data_loaded' not in st.session_state:
    st.session_state.data_loaded = False
if 'students_df' not in st.session_state:
    st.session_state.students_df = None
if 'companies_df' not in st.session_state:
    st.session_state.companies_df = None
if 'rankings_df' not in st.session_state:
    st.session_state.rankings_df = None

# Sidebar navigation
st.sidebar.title("🎓 Co-op Placement Optimizer")
page = st.sidebar.radio("Navigation", 
                        ["Data Setup", "Manual Data Editor", "Exploratory Analysis", "Clustering", "Optimization"],
                        index=0)

# Helper functions
def generate_synthetic_data(n_students=15, n_companies=15):
    """Generate synthetic dataset with flexible numbers"""
    np.random.seed(42)
    
    # Students
    student_names = [f"Student_{i+1:02d}" for i in range(n_students)]
    students_df = pd.DataFrame({
        'student_id': range(1, n_students+1),
        'student_name': student_names,
        'gpa': np.random.uniform(3.0, 4.0, n_students).round(2)
    })
    
    # Companies
    base_companies = [
        ("QBE", "General Insurance"),
        ("IAG", "General Insurance"),
        ("Suncorp", "General Insurance"),
        ("Allianz", "General Insurance"),
        ("Deloitte", "Consultancy"),
        ("PwC", "Consultancy"),
        ("KPMG", "Consultancy"),
        ("EY", "Consultancy"),
        ("AMP", "Life Insurance"),
        ("MLC", "Life Insurance"),
        ("TAL", "Life Insurance"),
        ("Zurich", "Life Insurance"),
        ("NDIS Provider A", "Care/Disability"),
        ("NDIS Provider B", "Care/Disability"),
        ("NDIS Provider C", "Care/Disability"),
    ]
    
    company_info = []
    industries = ['General Insurance', 'Consultancy', 'Life Insurance', 'Care/Disability']
    
    for i in range(n_companies):
        if i < len(base_companies):
            name, industry = base_companies[i]
        else:
            industry = industries[i % len(industries)]
            name = f"Company_{i+1:02d}"
        
        base_capacity = max(1, int(np.ceil(n_students * 1.2 / n_companies)))
        company_info.append((name, industry, base_capacity, base_capacity))
    
    companies_df = pd.DataFrame(company_info, 
                                columns=['company_name', 'industry', 'it2_capacity', 'it3_capacity'])
    companies_df['company_id'] = range(1, len(companies_df)+1)
    companies_df = companies_df[['company_id', 'company_name', 'industry', 'it2_capacity', 'it3_capacity']]
    
    # Rankings (higher is better)
    rankings_data = []
    for student_id in range(1, n_students+1):
        student_rankings = np.random.randint(1, 11, len(companies_df))
        for company_id, ranking in enumerate(student_rankings, 1):
            rankings_data.append({
                'student_id': student_id,
                'company_id': company_id,
                'ranking': ranking
            })
    
    rankings_df = pd.DataFrame(rankings_data)
    
    return students_df, companies_df, rankings_df

def initialize_empty_data():
    """Initialize empty dataframes"""
    students_df = pd.DataFrame(columns=['student_id', 'student_name', 'gpa'])
    companies_df = pd.DataFrame(columns=['company_id', 'company_name', 'industry', 'it2_capacity', 'it3_capacity'])
    rankings_df = pd.DataFrame(columns=['student_id', 'company_id', 'ranking'])
    return students_df, companies_df, rankings_df

def regenerate_rankings(students_df, companies_df):
    """Generate empty rankings for all student-company pairs"""
    rankings_data = []
    for student_id in students_df['student_id']:
        for company_id in companies_df['company_id']:
            rankings_data.append({
                'student_id': student_id,
                'company_id': company_id,
                'ranking': 5  # Default mid-range ranking
            })
    return pd.DataFrame(rankings_data)

def create_excel_template(n_students=20, n_companies=20):
    """Create Excel template for data input with flexible size"""
    wb = openpyxl.Workbook()
    
    # Students sheet
    ws_students = wb.active
    ws_students.title = "Students"
    ws_students['A1'] = 'student_id'
    ws_students['B1'] = 'student_name'
    ws_students['C1'] = 'gpa'
    
    for cell in ['A1', 'B1', 'C1']:
        ws_students[cell].font = Font(bold=True, color='FFFFFF')
        ws_students[cell].fill = PatternFill(start_color='366092', fill_type='solid')
    
    for i in range(2, n_students + 2):
        ws_students[f'A{i}'] = i-1
        ws_students[f'B{i}'] = f'Student_{i-1:02d}'
        ws_students[f'C{i}'] = 3.5
    
    # Companies sheet
    ws_companies = wb.create_sheet("Companies")
    headers = ['company_id', 'company_name', 'industry', 'it2_capacity', 'it3_capacity']
    for col, header in enumerate(headers, 1):
        cell = ws_companies.cell(1, col, header)
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill(start_color='366092', fill_type='solid')
    
    industries = ['General Insurance', 'Consultancy', 'Life Insurance', 'Care/Disability']
    for i in range(2, n_companies + 2):
        ws_companies.cell(i, 1, i-1)
        ws_companies.cell(i, 2, f'Company_{i-1:02d}')
        ws_companies.cell(i, 3, industries[(i-2) % 4])
        ws_companies.cell(i, 4, 1)
        ws_companies.cell(i, 5, 1)
    
    # Rankings sheet
    ws_rankings = wb.create_sheet("Rankings")
    ws_rankings['A1'] = 'student_id'
    ws_rankings['B1'] = 'company_id'
    ws_rankings['C1'] = 'ranking'
    
    for cell in ['A1', 'B1', 'C1']:
        ws_rankings[cell].font = Font(bold=True, color='FFFFFF')
        ws_rankings[cell].fill = PatternFill(start_color='366092', fill_type='solid')
    
    ws_rankings['E1'] = 'Instructions:'
    ws_rankings['E1'].font = Font(bold=True, size=12)
    ws_rankings['E2'] = '1. Fill in student information in Students sheet'
    ws_rankings['E3'] = '2. Fill in company information in Companies sheet'
    ws_rankings['E4'] = '3. Enter rankings (1-10, higher is better) for EACH student-company pair'
    ws_rankings['E5'] = '4. Rankings: Each row = one student ranking one company'
    ws_rankings['E6'] = '5. Total rows needed = (# students) × (# companies)'
    ws_rankings['E7'] = f'6. For this template: {n_students} students × {n_companies} companies = {n_students * n_companies} rankings'
    
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

def load_excel_data(file):
    """Load data from uploaded Excel file"""
    try:
        students_df = pd.read_excel(file, sheet_name='Students')
        companies_df = pd.read_excel(file, sheet_name='Companies')
        rankings_df = pd.read_excel(file, sheet_name='Rankings')
        return students_df, companies_df, rankings_df, None
    except Exception as e:
        return None, None, None, str(e)

def validate_data(students_df, companies_df, rankings_df):
    """Validate uploaded data"""
    errors = []
    warnings = []
    
    n_students = len(students_df)
    n_companies = len(companies_df)
    
    if n_students == 0:
        errors.append("No students found")
    if n_companies == 0:
        errors.append("No companies found")
    
    if n_students > 0 and n_companies > 0:
        expected_rankings = n_students * n_companies
        if len(rankings_df) != expected_rankings:
            errors.append(f"Expected {expected_rankings} rankings ({n_students} students × {n_companies} companies), got {len(rankings_df)}")
        
        if len(rankings_df) > 0:
            if rankings_df['ranking'].min() < 1 or rankings_df['ranking'].max() > 10:
                errors.append("Rankings must be between 1 and 10")
        
        total_it2_capacity = companies_df['it2_capacity'].sum()
        total_it3_capacity = companies_df['it3_capacity'].sum()
        
        if total_it2_capacity < n_students:
            errors.append(f"Total IT2 capacity ({total_it2_capacity}) is less than number of students ({n_students})")
        
        if total_it3_capacity < n_students:
            errors.append(f"Total IT3 capacity ({total_it3_capacity}) is less than number of students ({n_students})")
        
        if total_it2_capacity < n_students * 1.2:
            warnings.append(f"IT2 capacity is tight. Consider increasing for better optimization.")
        
        if total_it3_capacity < n_students * 1.2:
            warnings.append(f"IT3 capacity is tight. Consider increasing for better optimization.")
    
    return errors, warnings

# PAGE 1: DATA SETUP (Quick Start)
if page == "Data Setup":
    st.title("📊 Data Setup - Quick Start")
    
    tab1, tab2 = st.tabs(["Synthetic Data", "Excel Upload"])
    
    # Tab 1: Synthetic Data
    with tab1:
        st.header("Generate Synthetic Dataset")
        st.write("Generate test data with custom numbers of students and companies.")
        
        col1, col2 = st.columns(2)
        with col1:
            n_students = st.number_input("Number of Students", min_value=5, max_value=100, value=15, step=1)
        with col2:
            n_companies = st.number_input("Number of Companies", min_value=5, max_value=100, value=15, step=1)
        
        if st.button("Generate Synthetic Data", type="primary"):
            students_df, companies_df, rankings_df = generate_synthetic_data(n_students, n_companies)
            st.session_state.students_df = students_df
            st.session_state.companies_df = companies_df
            st.session_state.rankings_df = rankings_df
            st.session_state.data_loaded = True
            st.success(f"✅ Generated: {n_students} students, {n_companies} companies, {len(rankings_df)} rankings")
            st.info("💡 Go to 'Manual Data Editor' page to customize student names, companies, and rankings!")
        
        if st.session_state.data_loaded:
            st.subheader("Preview Data")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.write("**Students:**")
                st.dataframe(st.session_state.students_df, height=300)
            with col2:
                st.write("**Companies:**")
                st.dataframe(st.session_state.companies_df, height=300)
            with col3:
                st.write("**Rankings (sample):**")
                st.dataframe(st.session_state.rankings_df.head(20), height=300)
    
    # Tab 2: Excel Upload
    with tab2:
        st.header("Excel Data Upload")
        
        st.subheader("Step 1: Download Template")
        st.write("Customize template size based on your needs:")
        
        col1, col2 = st.columns(2)
        with col1:
            template_students = st.number_input("Students in template", min_value=5, max_value=200, value=20, step=5)
        with col2:
            template_companies = st.number_input("Companies in template", min_value=5, max_value=200, value=20, step=5)
        
        template_buffer = create_excel_template(template_students, template_companies)
        st.download_button(
            label=f"📥 Download Excel Template ({template_students} students × {template_companies} companies)",
            data=template_buffer,
            file_name=f"coop_placement_template_{template_students}x{template_companies}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.info(f"💡 This template has {template_students} students and {template_companies} companies. You can add/remove rows as needed in Excel.")
        
        st.subheader("Step 2: Upload Completed File")
        uploaded_file = st.file_uploader("Upload Excel file", type=['xlsx'])
        
        if uploaded_file:
            students_df, companies_df, rankings_df, error = load_excel_data(uploaded_file)
            
            if error:
                st.error(f"❌ Error loading file: {error}")
            else:
                errors, warnings = validate_data(students_df, companies_df, rankings_df)
                
                if errors:
                    st.error("❌ Data validation failed:")
                    for err in errors:
                        st.write(f"- {err}")
                else:
                    if warnings:
                        st.warning("⚠️ Warnings:")
                        for warn in warnings:
                            st.write(f"- {warn}")
                    
                    st.session_state.students_df = students_df
                    st.session_state.companies_df = companies_df
                    st.session_state.rankings_df = rankings_df
                    st.session_state.data_loaded = True
                    st.success("✅ Data loaded and validated successfully!")
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Students", len(students_df))
                    with col2:
                        st.metric("Companies", len(companies_df))
                    with col3:
                        st.metric("Rankings", len(rankings_df))

# PAGE 2: MANUAL DATA EDITOR (NEW!)
# =============================================================================
# UPDATED MANUAL DATA EDITOR SECTION - COMPLETELY EMPTY, NO PRE-FILLED DATA
# =============================================================================
# Replace the existing "Manual Data Editor" section (around line 322-636) 
# with this updated version

# PAGE 2: MANUAL DATA EDITOR (COMPLETELY EMPTY - NO PRE-FILLED DATA)
elif page == "Manual Data Editor":
    st.title("✏️ Manual Data Editor")
    st.write("Create and edit your data manually from scratch - **starts completely empty**. You fill in everything!")
    
    # Initialize completely empty data if not loaded
    if not st.session_state.data_loaded:
        st.session_state.students_df = pd.DataFrame(columns=['student_id', 'student_name', 'gpa'])
        st.session_state.companies_df = pd.DataFrame(columns=['company_id', 'company_name', 'industry', 'it2_capacity', 'it3_capacity'])
        st.session_state.rankings_df = pd.DataFrame(columns=['student_id', 'company_id', 'ranking'])
        st.session_state.data_loaded = True
    
    # Create tabs for different editors
    tab1, tab2, tab3, tab4 = st.tabs(["👥 Students", "🏢 Companies", "⭐ Rankings", "📊 Summary"])
    
    # TAB 1: STUDENTS EDITOR
    with tab1:
        st.header("Students Management")
        st.caption("⚠️ All fields start empty - you must fill in all information")
        
        # Add new student
        st.subheader("➕ Add New Student")
        col1, col2, col3 = st.columns([2, 2, 1])
        with col1:
            new_student_name = st.text_input("Student Name (required)", key="new_student_name", placeholder="Enter student name...")
        with col2:
            new_student_gpa = st.number_input("GPA", min_value=0.0, max_value=4.0, value=0.0, step=0.01, key="new_student_gpa", help="Enter GPA (0.0 to 4.0)")
        with col3:
            st.write("")
            st.write("")
            if st.button("Add Student", type="primary"):
                if new_student_name.strip():
                    new_id = st.session_state.students_df['student_id'].max() + 1 if len(st.session_state.students_df) > 0 else 1
                    new_row = pd.DataFrame({
                        'student_id': [new_id],
                        'student_name': [new_student_name.strip()],
                        'gpa': [new_student_gpa]
                    })
                    st.session_state.students_df = pd.concat([st.session_state.students_df, new_row], ignore_index=True)
                    
                    # Add EMPTY rankings for this new student with all companies (ranking = 0)
                    if len(st.session_state.companies_df) > 0:
                        new_rankings = []
                        for company_id in st.session_state.companies_df['company_id']:
                            new_rankings.append({
                                'student_id': new_id,
                                'company_id': company_id,
                                'ranking': 0  # Empty/unranked
                            })
                        st.session_state.rankings_df = pd.concat([st.session_state.rankings_df, pd.DataFrame(new_rankings)], ignore_index=True)
                    
                    st.success(f"✅ Added student: {new_student_name}")
                    st.rerun()
                else:
                    st.error("❌ Please enter a student name")
        
        # Display and edit existing students
        st.subheader("📋 Current Students")
        if len(st.session_state.students_df) > 0:
            edited_students = st.data_editor(
                st.session_state.students_df,
                use_container_width=True,
                num_rows="dynamic",
                column_config={
                    "student_id": st.column_config.NumberColumn("ID", disabled=True),
                    "student_name": st.column_config.TextColumn("Name", required=True),
                    "gpa": st.column_config.NumberColumn("GPA", min_value=0.0, max_value=4.0, step=0.01, format="%.2f")
                },
                hide_index=True,
                key="students_editor"
            )
            
            if st.button("💾 Save Students Changes"):
                st.session_state.students_df = edited_students
                # Reassign IDs
                st.session_state.students_df['student_id'] = range(1, len(st.session_state.students_df) + 1)
                st.success("✅ Students updated!")
                st.rerun()
        else:
            st.info("👆 No students yet. Add your first student above!")
    
    # TAB 2: COMPANIES EDITOR
    with tab2:
        st.header("Companies Management")
        st.caption("⚠️ All fields start empty - you must fill in all information")
        
        # Add new company
        st.subheader("➕ Add New Company")
        col1, col2, col3, col4 = st.columns([2, 2, 1, 1])
        with col1:
            new_company_name = st.text_input("Company Name (required)", key="new_company_name", placeholder="Enter company name...")
        with col2:
            new_company_industry = st.selectbox(
                "Industry",
                ["", "General Insurance", "Consultancy", "Life Insurance", "Care/Disability", "Banking", "Superannuation", "Other"],
                key="new_company_industry",
                help="Select industry type"
            )
        with col3:
            new_it2_cap = st.number_input("IT2 Cap", min_value=0, value=0, step=1, key="new_it2_cap", help="IT2 capacity")
        with col4:
            new_it3_cap = st.number_input("IT3 Cap", min_value=0, value=0, step=1, key="new_it3_cap", help="IT3 capacity")
        
        if st.button("Add Company", type="primary", key="add_company_btn"):
            if new_company_name.strip() and new_company_industry:
                new_id = st.session_state.companies_df['company_id'].max() + 1 if len(st.session_state.companies_df) > 0 else 1
                new_row = pd.DataFrame({
                    'company_id': [new_id],
                    'company_name': [new_company_name.strip()],
                    'industry': [new_company_industry],
                    'it2_capacity': [new_it2_cap],
                    'it3_capacity': [new_it3_cap]
                })
                st.session_state.companies_df = pd.concat([st.session_state.companies_df, new_row], ignore_index=True)
                
                # Add EMPTY rankings for all students with this new company (ranking = 0)
                if len(st.session_state.students_df) > 0:
                    new_rankings = []
                    for student_id in st.session_state.students_df['student_id']:
                        new_rankings.append({
                            'student_id': student_id,
                            'company_id': new_id,
                            'ranking': 0  # Empty/unranked
                        })
                    st.session_state.rankings_df = pd.concat([st.session_state.rankings_df, pd.DataFrame(new_rankings)], ignore_index=True)
                
                st.success(f"✅ Added company: {new_company_name}")
                st.rerun()
            else:
                st.error("❌ Please enter a company name and select an industry")
        
        # Display and edit existing companies
        st.subheader("📋 Current Companies")
        if len(st.session_state.companies_df) > 0:
            edited_companies = st.data_editor(
                st.session_state.companies_df,
                use_container_width=True,
                num_rows="dynamic",
                column_config={
                    "company_id": st.column_config.NumberColumn("ID", disabled=True),
                    "company_name": st.column_config.TextColumn("Company Name", required=True),
                    "industry": st.column_config.SelectboxColumn(
                        "Industry",
                        options=["General Insurance", "Consultancy", "Life Insurance", "Care/Disability", "Banking", "Superannuation", "Other"],
                        required=True
                    ),
                    "it2_capacity": st.column_config.NumberColumn("IT2 Capacity", min_value=0, step=1, format="%d"),
                    "it3_capacity": st.column_config.NumberColumn("IT3 Capacity", min_value=0, step=1, format="%d")
                },
                hide_index=True,
                key="companies_editor"
            )
            
            if st.button("💾 Save Companies Changes"):
                st.session_state.companies_df = edited_companies
                # Reassign IDs
                st.session_state.companies_df['company_id'] = range(1, len(st.session_state.companies_df) + 1)
                st.success("✅ Companies updated!")
                st.rerun()
        else:
            st.info("👆 No companies yet. Add your first company above!")
    
    # TAB 3: RANKINGS EDITOR
    with tab3:
        st.header("Rankings Management")
        
        if len(st.session_state.students_df) == 0:
            st.warning("⚠️ Please add students first in the Students tab")
        elif len(st.session_state.companies_df) == 0:
            st.warning("⚠️ Please add companies first in the Companies tab")
        else:
            # Initialize empty rankings structure if needed
            expected_rankings = len(st.session_state.students_df) * len(st.session_state.companies_df)
            if len(st.session_state.rankings_df) != expected_rankings:
                st.info("🔄 Generating empty ranking slots for all student-company pairs...")
                rankings_data = []
                for student_id in st.session_state.students_df['student_id']:
                    for company_id in st.session_state.companies_df['company_id']:
                        rankings_data.append({
                            'student_id': student_id,
                            'company_id': company_id,
                            'ranking': 0  # Start with 0 (empty/unranked)
                        })
                st.session_state.rankings_df = pd.DataFrame(rankings_data)
                st.rerun()
            
            st.caption("⚠️ All rankings start at 0 (unranked). You must fill in values from 1-10 for each student-company pair.")
            st.info(f"📊 Total rankings to fill: {len(st.session_state.students_df)} students × {len(st.session_state.companies_df)} companies = {expected_rankings} rankings")
            
            # Quick edit interface - by student
            st.subheader("✏️ Edit Rankings by Student")
            
            student_names = st.session_state.students_df.set_index('student_id')['student_name'].to_dict()
            company_names = st.session_state.companies_df.set_index('company_id')['company_name'].to_dict()
            
            selected_student_id = st.selectbox(
                "Select Student",
                st.session_state.students_df['student_id'].tolist(),
                format_func=lambda x: f"{x} - {student_names[x]}"
            )
            
            st.write(f"**Fill rankings for: {student_names[selected_student_id]}**")
            st.caption("0 = unranked, 1 = lowest preference, 10 = highest preference")
            
            # Get rankings for selected student
            student_rankings = st.session_state.rankings_df[
                st.session_state.rankings_df['student_id'] == selected_student_id
            ].copy()
            
            # Create editable view
            student_rankings['company_name'] = student_rankings['company_id'].map(company_names)
            student_rankings = student_rankings[['company_id', 'company_name', 'ranking']]
            
            edited_rankings = st.data_editor(
                student_rankings,
                use_container_width=True,
                column_config={
                    "company_id": st.column_config.NumberColumn("ID", disabled=True),
                    "company_name": st.column_config.TextColumn("Company", disabled=True),
                    "ranking": st.column_config.NumberColumn(
                        "Ranking (0-10) ⭐",
                        min_value=0,
                        max_value=10,
                        step=1,
                        help="0=Unranked, 1=Least preferred, 10=Most preferred"
                    )
                },
                hide_index=True,
                key=f"rankings_editor_{selected_student_id}"
            )
            
            if st.button("💾 Save Rankings for This Student"):
                # Update rankings in main dataframe
                for _, row in edited_rankings.iterrows():
                    mask = (st.session_state.rankings_df['student_id'] == selected_student_id) & \
                           (st.session_state.rankings_df['company_id'] == row['company_id'])
                    st.session_state.rankings_df.loc[mask, 'ranking'] = row['ranking']
                st.success(f"✅ Updated rankings for {student_names[selected_student_id]}")
                st.rerun()
            
            # Show progress
            st.markdown("---")
            st.subheader("📊 Progress")
            col1, col2, col3 = st.columns(3)
            with col1:
                filled = len(st.session_state.rankings_df[st.session_state.rankings_df['ranking'] > 0])
                st.metric("Filled Rankings", f"{filled} / {expected_rankings}")
            with col2:
                completion = (filled / expected_rankings * 100) if expected_rankings > 0 else 0
                st.metric("Completion", f"{completion:.1f}%")
            with col3:
                if filled > 0:
                    avg_rank = st.session_state.rankings_df[st.session_state.rankings_df['ranking'] > 0]['ranking'].mean()
                    st.metric("Average Ranking", f"{avg_rank:.2f}")
                else:
                    st.metric("Average Ranking", "N/A")
            
            # Show all rankings table
            st.subheader("📋 All Rankings (Overview)")
            rankings_display = st.session_state.rankings_df.copy()
            rankings_display['student_name'] = rankings_display['student_id'].map(student_names)
            rankings_display['company_name'] = rankings_display['company_id'].map(company_names)
            rankings_display = rankings_display[['student_name', 'company_name', 'ranking']]
            
            st.dataframe(
                rankings_display,
                use_container_width=True,
                height=400,
                hide_index=True
            )
    
    # TAB 4: SUMMARY
    with tab4:
        st.header("📊 Data Summary")
        
        if st.session_state.data_loaded and (len(st.session_state.students_df) > 0 or len(st.session_state.companies_df) > 0):
            errors, warnings = validate_data(
                st.session_state.students_df,
                st.session_state.companies_df,
                st.session_state.rankings_df
            )
            
            # Metrics
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Students", len(st.session_state.students_df))
            with col2:
                st.metric("Companies", len(st.session_state.companies_df))
            with col3:
                total_rankings = len(st.session_state.rankings_df)
                filled_rankings = len(st.session_state.rankings_df[st.session_state.rankings_df['ranking'] > 0])
                st.metric("Rankings Filled", f"{filled_rankings} / {total_rankings}")
            with col4:
                expected = len(st.session_state.students_df) * len(st.session_state.companies_df)
                completion = (filled_rankings / expected * 100) if expected > 0 else 0
                st.metric("Completion", f"{completion:.0f}%")
            
            # Validation
            st.subheader("✅ Validation")
            if errors:
                st.error("❌ Errors Found:")
                for err in errors:
                    st.write(f"- {err}")
            else:
                st.success("✅ All validation checks passed!")
            
            if warnings:
                st.warning("⚠️ Warnings:")
                for warn in warnings:
                    st.write(f"- {warn}")
            
            # Capacity analysis
            st.subheader("📊 Capacity Analysis")
            if len(st.session_state.companies_df) > 0:
                col1, col2 = st.columns(2)
                with col1:
                    total_it2 = st.session_state.companies_df['it2_capacity'].sum()
                    st.metric("Total IT2 Capacity", total_it2)
                    if len(st.session_state.students_df) > 0:
                        buffer = total_it2 - len(st.session_state.students_df)
                        st.write(f"Buffer: {buffer} spots")
                with col2:
                    total_it3 = st.session_state.companies_df['it3_capacity'].sum()
                    st.metric("Total IT3 Capacity", total_it3)
                    if len(st.session_state.students_df) > 0:
                        buffer = total_it3 - len(st.session_state.students_df)
                        st.write(f"Buffer: {buffer} spots")
            
            # Download option
            st.markdown("---")
            st.subheader("💾 Export Your Data")
            if len(st.session_state.students_df) > 0 and len(st.session_state.companies_df) > 0:
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    st.session_state.students_df.to_excel(writer, sheet_name='Students', index=False)
                    st.session_state.companies_df.to_excel(writer, sheet_name='Companies', index=False)
                    st.session_state.rankings_df.to_excel(writer, sheet_name='Rankings', index=False)
                
                st.download_button(
                    label="📥 Download as Excel",
                    data=output.getvalue(),
                    file_name="coop_data_manual.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            # Ready to optimize?
            st.markdown("---")
            st.subheader("🚀 Ready to Optimize?")
            if not errors and filled_rankings == expected:
                st.success("✅ Your data is complete and ready! Go to the 'Optimization' page to solve.")
            elif not errors and filled_rankings < expected:
                st.warning(f"⚠️ Data is valid but only {completion:.0f}% of rankings are filled. Fill all rankings for best results.")
            else:
                st.error("❌ Fix the errors above before optimizing")
            
            # Clear all button
            st.markdown("---")
            if st.button("🗑️ Clear All Data", type="secondary"):
                st.session_state.students_df = pd.DataFrame(columns=['student_id', 'student_name', 'gpa'])
                st.session_state.companies_df = pd.DataFrame(columns=['company_id', 'company_name', 'industry', 'it2_capacity', 'it3_capacity'])
                st.session_state.rankings_df = pd.DataFrame(columns=['student_id', 'company_id', 'ranking'])
                st.rerun()
        else:
            st.info("💡 No data yet. Use the tabs above to add students, companies, and rankings from scratch.")
# PAGE 3: EXPLORATORY ANALYSIS (Same as before, just renumbered)
elif page == "Exploratory Analysis":
    st.title("📈 Exploratory Data Analysis")
    
    if not st.session_state.data_loaded:
        st.warning("⚠️ Please load data first in the Data Setup page.")
    else:
        students_df = st.session_state.students_df
        companies_df = st.session_state.companies_df
        rankings_df = st.session_state.rankings_df
        
        # Summary statistics
        st.header("Summary Statistics")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Students", len(students_df))
        with col2:
            st.metric("Total Companies", len(companies_df))
        with col3:
            st.metric("Average Ranking", f"{rankings_df['ranking'].mean():.2f}")
        with col4:
            st.metric("Total IT2 Capacity", companies_df['it2_capacity'].sum())
        
        # Industry distribution
        st.header("Industry Distribution")
        col1, col2 = st.columns(2)
        
        with col1:
            industry_counts = companies_df['industry'].value_counts()
            fig = px.pie(values=industry_counts.values, names=industry_counts.index,
                        title="Companies by Industry")
            st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            st.write("**Industry Breakdown:**")
            st.dataframe(industry_counts.to_frame('Count'), height=200)
        
        # Ranking distribution
        st.header("Ranking Distribution")
        col1, col2 = st.columns(2)
        
        with col1:
            fig = px.histogram(rankings_df, x='ranking', nbins=10,
                              title="Distribution of Rankings",
                              labels={'ranking': 'Ranking', 'count': 'Frequency'})
            st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            fig = px.box(rankings_df, y='ranking', title="Ranking Box Plot")
            st.plotly_chart(fig, use_container_width=True)
        
        # Average ranking by company
        st.header("Average Rankings by Company")
        avg_rankings = rankings_df.groupby('company_id')['ranking'].mean().reset_index()
        avg_rankings = avg_rankings.merge(companies_df[['company_id', 'company_name', 'industry']], on='company_id')
        avg_rankings = avg_rankings.sort_values('ranking', ascending=False)
        
        fig = px.bar(avg_rankings.head(20), x='company_name', y='ranking', color='industry',
                    title="Top 20 Companies by Average Ranking",
                    labels={'ranking': 'Average Ranking', 'company_name': 'Company'})
        fig.update_xaxes(tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)
        
        # Heatmap
        st.header("Student-Company Ranking Heatmap")
        if len(students_df) <= 30 and len(companies_df) <= 30:
            pivot_rankings = rankings_df.pivot(index='student_id', columns='company_id', values='ranking')
            
            fig = px.imshow(pivot_rankings, 
                           labels=dict(x="Company ID", y="Student ID", color="Ranking"),
                           title="Student Preferences Heatmap",
                           aspect="auto",
                           color_continuous_scale='RdYlGn')
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info(f"Heatmap hidden for large datasets ({len(students_df)} students × {len(companies_df)} companies).")

# PAGE 4: CLUSTERING (Same as before, just renumbered)
elif page == "Clustering":
    st.title("🔗 Hierarchical Clustering")
    
    if not st.session_state.data_loaded:
        st.warning("⚠️ Please load data first in the Data Setup page.")
    else:
        st.write("""
        This page performs hierarchical clustering on students based on their company preferences.
        Students with similar ranking patterns will be grouped together.
        """)
        
        students_df = st.session_state.students_df
        rankings_df = st.session_state.rankings_df
        
        if len(students_df) > 100:
            st.warning("⚠️ Clustering may be slow for large datasets (>100 students)")
        
        pivot_rankings = rankings_df.pivot(index='student_id', columns='company_id', values='ranking')
        
        linkage_methods = ['ward', 'complete', 'average', 'single']
        method = st.selectbox("Select Linkage Method", linkage_methods, index=0)
        
        Z = linkage(pivot_rankings, method=method)
        
        st.header("Dendrogram")
        fig = go.Figure()
        
        from scipy.cluster.hierarchy import dendrogram as scipy_dendrogram
        dend = scipy_dendrogram(Z, no_plot=True)
        
        icoord = np.array(dend['icoord'])
        dcoord = np.array(dend['dcoord'])
        
        for i in range(len(icoord)):
            fig.add_trace(go.Scatter(
                x=icoord[i], y=dcoord[i],
                mode='lines',
                line=dict(color='rgb(100,100,100)', width=1),
                hoverinfo='skip',
                showlegend=False
            ))
        
        fig.update_layout(
            title=f"Hierarchical Clustering Dendrogram (Method: {method})",
            xaxis_title="Student ID",
            yaxis_title="Distance",
            height=500
        )
        st.plotly_chart(fig, use_container_width=True)
        
        st.header("Cluster Assignment")
        max_clusters = min(10, len(students_df) - 1)
        n_clusters = st.slider("Select Number of Clusters", 2, max_clusters, min(3, max_clusters))
        
        clusters = fcluster(Z, n_clusters, criterion='maxclust')
        
        students_clustered = students_df.copy()
        students_clustered['cluster'] = clusters
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Cluster Sizes")
            cluster_sizes = students_clustered['cluster'].value_counts().sort_index()
            fig = px.bar(x=cluster_sizes.index, y=cluster_sizes.values,
                        labels={'x': 'Cluster', 'y': 'Number of Students'},
                        title=f"Students per Cluster (Total: {n_clusters} clusters)")
            st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            st.subheader("Cluster Assignments")
            st.dataframe(students_clustered[['student_name', 'cluster']].sort_values('cluster'),
                        height=400)
        
        st.header("Cluster Characteristics")
        for cluster_id in range(1, n_clusters + 1):
            with st.expander(f"Cluster {cluster_id} ({cluster_sizes[cluster_id]} students)"):
                cluster_students = students_clustered[students_clustered['cluster'] == cluster_id]['student_id'].tolist()
                cluster_rankings = rankings_df[rankings_df['student_id'].isin(cluster_students)]
                
                avg_pref = cluster_rankings.groupby('company_id')['ranking'].mean().sort_values(ascending=False).head(5)
                avg_pref = avg_pref.reset_index()
                avg_pref = avg_pref.merge(st.session_state.companies_df[['company_id', 'company_name']], on='company_id')
                
                st.write("**Top 5 Preferred Companies (average):**")
                st.dataframe(avg_pref[['company_name', 'ranking']], hide_index=True)

# PAGE 5: OPTIMIZATION (Same as before, just renumbered)
elif page == "Optimization":
    st.title("⚡ Placement Optimization")
    
    if not st.session_state.data_loaded:
        st.warning("⚠️ Please load data first in the Data Setup page.")
    else:
        st.write("""
        This page solves the integer linear programming problem to optimally assign students 
        to companies for IT2 and IT3 placements while respecting all constraints.
        """)
        
        students_df = st.session_state.students_df
        companies_df = st.session_state.companies_df
        rankings_df = st.session_state.rankings_df
        
        n_students = len(students_df)
        n_companies = len(companies_df)
        n_variables = n_students * n_companies * 2
        
        st.info(f"**Problem Size:** {n_students} students × {n_companies} companies = {n_variables} decision variables")
        
        if n_variables > 10000:
            st.warning("⚠️ Large problem size. Optimization may take several minutes.")
        
        if st.button("🚀 Solve Optimization Problem", type="primary"):
            with st.spinner("Solving optimization problem..."):
                prob = pulp.LpProblem("CoopPlacement", pulp.LpMaximize)
                
                students = students_df['student_id'].tolist()
                companies = companies_df['company_id'].tolist()
                
                x = {}
                y = {}
                
                for i in students:
                    for j in companies:
                        x[i, j] = pulp.LpVariable(f"x_{i}_{j}", cat='Binary')
                        y[i, j] = pulp.LpVariable(f"y_{i}_{j}", cat='Binary')
                
                rankings_dict = rankings_df.set_index(['student_id', 'company_id'])['ranking'].to_dict()
                prob += pulp.lpSum([rankings_dict[i, j] * (x[i, j] + y[i, j]) 
                                   for i in students for j in companies])
                
                for i in students:
                    prob += pulp.lpSum([x[i, j] for j in companies]) == 1
                
                for i in students:
                    prob += pulp.lpSum([y[i, j] for j in companies]) == 1
                
                for j in companies:
                    capacity = companies_df[companies_df['company_id'] == j]['it2_capacity'].values[0]
                    prob += pulp.lpSum([x[i, j] for i in students]) <= capacity
                
                for j in companies:
                    capacity = companies_df[companies_df['company_id'] == j]['it3_capacity'].values[0]
                    prob += pulp.lpSum([y[i, j] for i in students]) <= capacity
                
                for i in students:
                    for j in companies:
                        prob += x[i, j] + y[i, j] <= 1
                
                industry_companies = companies_df.groupby('industry')['company_id'].apply(list).to_dict()
                
                for i in students:
                    for industry, company_list in industry_companies.items():
                        prob += pulp.lpSum([x[i, j] + y[i, j] for j in company_list]) <= 1
                
                solver = pulp.PULP_CBC_CMD(msg=0)
                prob.solve(solver)
                
                st.success(f"✅ Optimization Status: {pulp.LpStatus[prob.status]}")
                
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("Objective Value", f"{pulp.value(prob.objective):.2f}")
                with col2:
                    st.metric("Max Possible", f"{rankings_df['ranking'].max() * 2 * len(students)}")
                
                if prob.status == 1:
                    it2_assignments = []
                    it3_assignments = []
                    
                    for i in students:
                        for j in companies:
                            if pulp.value(x[i, j]) == 1:
                                student_name = students_df[students_df['student_id'] == i]['student_name'].values[0]
                                company_name = companies_df[companies_df['company_id'] == j]['company_name'].values[0]
                                industry = companies_df[companies_df['company_id'] == j]['industry'].values[0]
                                ranking = rankings_dict[i, j]
                                it2_assignments.append({
                                    'Student': student_name,
                                    'Company': company_name,
                                    'Industry': industry,
                                    'Ranking': ranking
                                })
                            
                            if pulp.value(y[i, j]) == 1:
                                student_name = students_df[students_df['student_id'] == i]['student_name'].values[0]
                                company_name = companies_df[companies_df['company_id'] == j]['company_name'].values[0]
                                industry = companies_df[companies_df['company_id'] == j]['industry'].values[0]
                                ranking = rankings_dict[i, j]
                                it3_assignments.append({
                                    'Student': student_name,
                                    'Company': company_name,
                                    'Industry': industry,
                                    'Ranking': ranking
                                })
                    
                    it2_df = pd.DataFrame(it2_assignments)
                    it3_df = pd.DataFrame(it3_assignments)
                    
                    st.header("Placement Assignments")
                    
                    tab1, tab2, tab3 = st.tabs(["IT2 Placements", "IT3 Placements", "Combined View"])
                    
                    with tab1:
                        st.subheader("IT2 Placements")
                        st.dataframe(it2_df.sort_values('Student'), hide_index=True, height=500)
                        st.metric("Average IT2 Ranking", f"{it2_df['Ranking'].mean():.2f}")
                    
                    with tab2:
                        st.subheader("IT3 Placements")
                        st.dataframe(it3_df.sort_values('Student'), hide_index=True, height=500)
                        st.metric("Average IT3 Ranking", f"{it3_df['Ranking'].mean():.2f}")
                    
                    with tab3:
                        combined = it2_df.merge(it3_df, on='Student', suffixes=('_IT2', '_IT3'))
                        st.dataframe(combined, hide_index=True, height=500)
                    
                    st.header("Placement Analysis")
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        fig = px.histogram(pd.concat([it2_df['Ranking'], it3_df['Ranking']]),
                                         x='Ranking', nbins=10,
                                         title="Distribution of Assigned Rankings")
                        st.plotly_chart(fig, use_container_width=True)
                    
                    with col2:
                        industry_dist = pd.concat([
                            it2_df['Industry'].value_counts().rename('IT2'),
                            it3_df['Industry'].value_counts().rename('IT3')
                        ], axis=1).fillna(0)
                        
                        fig = px.bar(industry_dist, barmode='group',
                                   title="Placements by Industry")
                        st.plotly_chart(fig, use_container_width=True)
                    
                    st.header("Export Results")
                    
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        it2_df.to_excel(writer, sheet_name='IT2_Placements', index=False)
                        it3_df.to_excel(writer, sheet_name='IT3_Placements', index=False)
                        combined.to_excel(writer, sheet_name='Combined', index=False)
                    
                    st.download_button(
                        label="📥 Download Results (Excel)",
                        data=output.getvalue(),
                        file_name="placement_results.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("❌ Optimization failed to find a solution. Check constraints.")
                    st.write("**Possible issues:**")
                    st.write("- Total company capacity < Number of students")
                    st.write("- Infeasible industry diversity constraints")
                    st.write("- Check your data for inconsistencies")

# Footer
st.sidebar.markdown("---")
st.sidebar.info(f"""
**Co-op Placement Optimizer**  
Version 3.0 - Manual Editor
Built with Streamlit & PuLP

Current Data:
- Students: {len(st.session_state.students_df) if st.session_state.data_loaded else 0}
- Companies: {len(st.session_state.companies_df) if st.session_state.data_loaded else 0}
""")
