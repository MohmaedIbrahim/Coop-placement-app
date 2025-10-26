import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from scipy.cluster.hierarchy import dendrogram, linkage, fcluster
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
        'student_name': student_names
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
    students_df = pd.DataFrame(columns=['student_id', 'student_name'])
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
    
    for cell in ['A1', 'B1']:
        ws_students[cell].font = Font(bold=True, color='FFFFFF')
        ws_students[cell].fill = PatternFill(start_color='366092', fill_type='solid')
    
    for i in range(2, n_students + 2):
        ws_students[f'A{i}'] = i-1
        ws_students[f'B{i}'] = f'Student_{i-1:02d}'
    
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
        ws_companies.cell(i, 4, 1)  # Default IT2 capacity
        ws_companies.cell(i, 5, 1)  # Default IT3 capacity
    
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
    ws_rankings['E4'] = '   Note: Capacities can be 0 (company not offering that iteration)'
    ws_rankings['E5'] = '3. Enter rankings (1-10, higher is better) for EACH student-company pair'
    ws_rankings['E6'] = '4. Rankings: Each row = one student ranking one company'
    ws_rankings['E7'] = '5. Total rows needed = (# students) × (# companies)'
    ws_rankings['E8'] = f'6. For this template: {n_students} students × {n_companies} companies = {n_students * n_companies} rankings'
    
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
    """Validate uploaded data with comprehensive checks"""
    errors = []
    warnings = []
    
    n_students = len(students_df)
    n_companies = len(companies_df)
    
    # Basic existence checks
    if n_students == 0:
        errors.append("No students found")
    if n_companies == 0:
        errors.append("No companies found")
    
    # Check for missing or null industry assignments
    if n_companies > 0:
        if companies_df['industry'].isna().any():
            null_companies = companies_df[companies_df['industry'].isna()]['company_name'].tolist()
            errors.append(f"Some companies have no industry assigned: {', '.join(map(str, null_companies))}")
        
        # Check if companies have empty string industries
        if (companies_df['industry'] == '').any():
            empty_companies = companies_df[companies_df['industry'] == '']['company_name'].tolist()
            errors.append(f"Some companies have empty industry: {', '.join(map(str, empty_companies))}")
    
    if n_students > 0 and n_companies > 0:
        # Rankings completeness check
        expected_rankings = n_students * n_companies
        if len(rankings_df) != expected_rankings:
            errors.append(f"Expected {expected_rankings} rankings ({n_students} students × {n_companies} companies), got {len(rankings_df)}")
        
        # Rankings range check
        if len(rankings_df) > 0:
            if rankings_df['ranking'].min() < 1 or rankings_df['ranking'].max() > 10:
                errors.append("Rankings must be between 1 and 10")
        
        # Capacity checks
        total_it2_capacity = companies_df['it2_capacity'].sum()
        total_it3_capacity = companies_df['it3_capacity'].sum()
        
        if total_it2_capacity < n_students:
            errors.append(f"Total IT2 capacity ({total_it2_capacity}) is less than number of students ({n_students})")
        
        if total_it3_capacity < n_students:
            errors.append(f"Total IT3 capacity ({total_it3_capacity}) is less than number of students ({n_students})")
        
        # Capacity warnings for tight constraints
        if total_it2_capacity < n_students * 1.2:
            warnings.append(f"IT2 capacity is tight ({total_it2_capacity} for {n_students} students). Consider increasing for better optimization.")
        
        if total_it3_capacity < n_students * 1.2:
            warnings.append(f"IT3 capacity is tight ({total_it3_capacity} for {n_students} students). Consider increasing for better optimization.")
        
        # Industry diversity feasibility check
        if n_companies > 0 and not companies_df['industry'].isna().all():
            industry_capacities = companies_df.groupby('industry').agg({
                'it2_capacity': 'sum',
                'it3_capacity': 'sum'
            })
            
            for industry, row in industry_capacities.iterrows():
                total_industry_capacity = row['it2_capacity'] + row['it3_capacity']
                if total_industry_capacity < n_students:
                    warnings.append(
                        f"Industry '{industry}' has limited capacity ({int(total_industry_capacity)} total slots). "
                        f"If many students prefer this industry, diversity constraints may be tight."
                    )
    
    return errors, warnings

def check_optimization_feasibility(students_df, companies_df):
    """
    Perform detailed feasibility checks before optimization
    Returns: (is_feasible, error_messages, warning_messages)
    """
    errors = []
    warnings = []
    
    n_students = len(students_df)
    total_it2_capacity = companies_df['it2_capacity'].sum()
    total_it3_capacity = companies_df['it3_capacity'].sum()
    
    # Critical capacity check
    if total_it2_capacity < n_students:
        errors.append(f"❌ INFEASIBLE: IT2 capacity ({total_it2_capacity}) < students ({n_students})")
    
    if total_it3_capacity < n_students:
        errors.append(f"❌ INFEASIBLE: IT3 capacity ({total_it3_capacity}) < students ({n_students})")
    
    # Industry diversity feasibility analysis
    industry_capacities = companies_df.groupby('industry').agg({
        'it2_capacity': 'sum',
        'it3_capacity': 'sum',
        'company_id': 'count'
    }).rename(columns={'company_id': 'num_companies'})
    
    for industry, row in industry_capacities.iterrows():
        total_capacity = row['it2_capacity'] + row['it3_capacity']
        
        # Each student can do at most 1 placement in each industry
        # So we need: total_industry_capacity >= n_students for full flexibility
        if total_capacity < n_students:
            warnings.append(
                f"⚠️ Industry '{industry}' has only {int(total_capacity)} total slots for {n_students} students. "
                f"Diversity constraints may force suboptimal assignments."
            )
        
        # Check if any industry dominates capacity
        total_capacity_all = total_it2_capacity + total_it3_capacity
        industry_percentage = (total_capacity / total_capacity_all) * 100
        
        if industry_percentage > 50:
            warnings.append(
                f"⚠️ Industry '{industry}' dominates capacity ({industry_percentage:.1f}% of total). "
                f"This may limit placement diversity."
            )
    
    # Check for industries with very low capacity
    min_recommended_capacity = n_students * 0.15  # At least 15% of students should be placeable
    for industry, row in industry_capacities.iterrows():
        total_capacity = row['it2_capacity'] + row['it3_capacity']
        if total_capacity < min_recommended_capacity:
            warnings.append(
                f"⚠️ Industry '{industry}' has very low capacity ({int(total_capacity)} slots). "
                f"Consider adding more companies or increasing capacity."
            )
    
    is_feasible = len(errors) == 0
    return is_feasible, errors, warnings

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
            n_students = st.number_input("Number of Students", min_value=1, max_value=100, value=15, step=1)
        with col2:
            n_companies = st.number_input("Number of Companies", min_value=1, max_value=100, value=15, step=1)
        
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
            template_students = st.number_input("Students in template", min_value=1, max_value=200, value=20, step=5)
        with col2:
            template_companies = st.number_input("Companies in template", min_value=1, max_value=200, value=20, step=5)
        
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

# PAGE 2: MANUAL DATA EDITOR
elif page == "Manual Data Editor":
    st.title("✏️ Manual Data Editor")
    st.write("Create and edit your data manually - full control over students, companies, and rankings.")
    
    # Initialize or start from scratch
    col1, col2 = st.columns([3, 1])
    with col1:
        if not st.session_state.data_loaded:
            st.info("💡 No data loaded. Choose an option below to start.")
    with col2:
        if st.button("🗑️ Clear All Data"):
            st.session_state.students_df, st.session_state.companies_df, st.session_state.rankings_df = initialize_empty_data()
            st.session_state.data_loaded = False
            st.rerun()
    
    # Create tabs for different editors
    tab1, tab2, tab3, tab4 = st.tabs(["👥 Students", "🏢 Companies", "⭐ Rankings", "📊 Summary"])
    
    # TAB 1: STUDENTS EDITOR
    with tab1:
        st.header("Students Management")
        
        if st.session_state.students_df is None or len(st.session_state.students_df) == 0:
            st.session_state.students_df = pd.DataFrame(columns=['student_id', 'student_name'])
        
        st.subheader("Add New Student")
        col1, col2 = st.columns([3, 1])
        with col1:
            new_student_name = st.text_input("Student Name", key="new_student_name")
        with col2:
            if st.button("➕ Add Student"):
                if new_student_name:
                    new_id = 1 if len(st.session_state.students_df) == 0 else st.session_state.students_df['student_id'].max() + 1
                    new_row = pd.DataFrame({'student_id': [new_id], 'student_name': [new_student_name]})
                    st.session_state.students_df = pd.concat([st.session_state.students_df, new_row], ignore_index=True)
                    st.success(f"Added {new_student_name}")
                    st.rerun()
                else:
                    st.warning("Please enter a student name")
        
        st.subheader("Current Students")
        if len(st.session_state.students_df) > 0:
            edited_students = st.data_editor(
                st.session_state.students_df,
                num_rows="dynamic",
                use_container_width=True,
                key="students_editor",
                hide_index=True
            )
            
            if st.button("💾 Save Student Changes"):
                st.session_state.students_df = edited_students
                # Regenerate rankings if companies exist
                if st.session_state.companies_df is not None and len(st.session_state.companies_df) > 0:
                    st.session_state.rankings_df = regenerate_rankings(
                        st.session_state.students_df,
                        st.session_state.companies_df
                    )
                    st.info("Rankings regenerated for all student-company pairs")
                st.session_state.data_loaded = True
                st.success("Student data saved!")
                st.rerun()
        else:
            st.info("No students yet. Add students above.")
    
    # TAB 2: COMPANIES EDITOR
    with tab2:
        st.header("Companies Management")
        
        if st.session_state.companies_df is None or len(st.session_state.companies_df) == 0:
            st.session_state.companies_df = pd.DataFrame(
                columns=['company_id', 'company_name', 'industry', 'it2_capacity', 'it3_capacity']
            )
        
        st.subheader("Add New Company")
        col1, col2, col3, col4 = st.columns([2, 2, 1, 1])
        with col1:
            new_company_name = st.text_input("Company Name", key="new_company_name")
        with col2:
            new_industry = st.selectbox("Industry", 
                                       ["General Insurance", "Consultancy", "Life Insurance", "Care/Disability", "Other"],
                                       key="new_industry")
        with col3:
            new_it2_cap = st.number_input("IT2 Cap", min_value=0, value=1, key="new_it2_cap")
        with col4:
            new_it3_cap = st.number_input("IT3 Cap", min_value=0, value=1, key="new_it3_cap")
        
        if st.button("➕ Add Company"):
            if new_company_name:
                new_id = 1 if len(st.session_state.companies_df) == 0 else st.session_state.companies_df['company_id'].max() + 1
                new_row = pd.DataFrame({
                    'company_id': [new_id],
                    'company_name': [new_company_name],
                    'industry': [new_industry],
                    'it2_capacity': [new_it2_cap],
                    'it3_capacity': [new_it3_cap]
                })
                st.session_state.companies_df = pd.concat([st.session_state.companies_df, new_row], ignore_index=True)
                st.success(f"Added {new_company_name}")
                st.rerun()
            else:
                st.warning("Please enter a company name")
        
        st.subheader("Current Companies")
        if len(st.session_state.companies_df) > 0:
            edited_companies = st.data_editor(
                st.session_state.companies_df,
                num_rows="dynamic",
                use_container_width=True,
                key="companies_editor",
                hide_index=True,
                column_config={
                    "industry": st.column_config.SelectboxColumn(
                        "Industry",
                        options=["General Insurance", "Consultancy", "Life Insurance", "Care/Disability", "Other"],
                        required=True
                    )
                }
            )
            
            if st.button("💾 Save Company Changes"):
                st.session_state.companies_df = edited_companies
                # Regenerate rankings if students exist
                if st.session_state.students_df is not None and len(st.session_state.students_df) > 0:
                    st.session_state.rankings_df = regenerate_rankings(
                        st.session_state.students_df,
                        st.session_state.companies_df
                    )
                    st.info("Rankings regenerated for all student-company pairs")
                st.session_state.data_loaded = True
                st.success("Company data saved!")
                st.rerun()
        else:
            st.info("No companies yet. Add companies above.")
    
    # TAB 3: RANKINGS EDITOR
    with tab3:
        st.header("Rankings Management")
        
        if st.session_state.students_df is None or len(st.session_state.students_df) == 0:
            st.warning("⚠️ Please add students first in the Students tab")
        elif st.session_state.companies_df is None or len(st.session_state.companies_df) == 0:
            st.warning("⚠️ Please add companies first in the Companies tab")
        else:
            # Initialize rankings if needed
            if st.session_state.rankings_df is None or len(st.session_state.rankings_df) == 0:
                st.session_state.rankings_df = regenerate_rankings(
                    st.session_state.students_df,
                    st.session_state.companies_df
                )
            
            st.write("**Edit Rankings (1-10, where 10 is highest preference)**")
            st.info(f"Total rankings to edit: {len(st.session_state.students_df)} students × {len(st.session_state.companies_df)} companies = {len(st.session_state.rankings_df)} rankings")
            
            # Create a pivot table for easier editing
            rankings_pivot = st.session_state.rankings_df.pivot(
                index='student_id',
                columns='company_id',
                values='ranking'
            )
            
            # Add student names as index
            student_names = st.session_state.students_df.set_index('student_id')['student_name']
            rankings_pivot.index = rankings_pivot.index.map(student_names)
            
            # Add company names as columns
            company_names = st.session_state.companies_df.set_index('company_id')['company_name']
            rankings_pivot.columns = rankings_pivot.columns.map(company_names)
            
            # Allow editing
            edited_rankings = st.data_editor(
                rankings_pivot,
                use_container_width=True,
                key="rankings_editor"
            )
            
            if st.button("💾 Save Rankings"):
                # Convert back to long format
                rankings_long = edited_rankings.reset_index().melt(
                    id_vars='index',
                    var_name='company_name',
                    value_name='ranking'
                )
                rankings_long.columns = ['student_name', 'company_name', 'ranking']
                
                # Map back to IDs
                student_id_map = st.session_state.students_df.set_index('student_name')['student_id']
                company_id_map = st.session_state.companies_df.set_index('company_name')['company_id']
                
                rankings_long['student_id'] = rankings_long['student_name'].map(student_id_map)
                rankings_long['company_id'] = rankings_long['company_name'].map(company_id_map)
                
                st.session_state.rankings_df = rankings_long[['student_id', 'company_id', 'ranking']]
                st.session_state.data_loaded = True
                st.success("Rankings saved!")
                st.rerun()
            
            # Bulk operations
            st.subheader("Bulk Operations")
            col1, col2 = st.columns(2)
            with col1:
                if st.button("🔄 Reset All Rankings to 5"):
                    st.session_state.rankings_df['ranking'] = 5
                    st.success("All rankings reset to 5")
                    st.rerun()
            with col2:
                if st.button("🎲 Randomize Rankings"):
                    st.session_state.rankings_df['ranking'] = np.random.randint(1, 11, len(st.session_state.rankings_df))
                    st.success("Rankings randomized")
                    st.rerun()
    
    # TAB 4: SUMMARY
    with tab4:
        st.header("Data Summary")
        
        if not st.session_state.data_loaded:
            st.warning("⚠️ No data loaded yet")
        else:
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Students", len(st.session_state.students_df))
            with col2:
                st.metric("Companies", len(st.session_state.companies_df))
            with col3:
                st.metric("Rankings", len(st.session_state.rankings_df))
            
            st.subheader("Data Validation")
            errors, warnings = validate_data(
                st.session_state.students_df,
                st.session_state.companies_df,
                st.session_state.rankings_df
            )
            
            if errors:
                st.error("❌ Data Issues Found:")
                for err in errors:
                    st.write(f"- {err}")
            else:
                st.success("✅ Data validation passed!")
            
            if warnings:
                st.warning("⚠️ Warnings:")
                for warn in warnings:
                    st.write(f"- {warn}")
            
            # Export functionality
            st.subheader("Export Data")
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                st.session_state.students_df.to_excel(writer, sheet_name='Students', index=False)
                st.session_state.companies_df.to_excel(writer, sheet_name='Companies', index=False)
                st.session_state.rankings_df.to_excel(writer, sheet_name='Rankings', index=False)
            
            st.download_button(
                label="📥 Download Data (Excel)",
                data=output.getvalue(),
                file_name="coop_placement_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# PAGE 3: EXPLORATORY ANALYSIS
elif page == "Exploratory Analysis":
    st.title("📊 Exploratory Data Analysis")
    
    if not st.session_state.data_loaded:
        st.warning("⚠️ Please load data first in the Data Setup page.")
    else:
        students_df = st.session_state.students_df
        companies_df = st.session_state.companies_df
        rankings_df = st.session_state.rankings_df
        
        # Overview metrics
        st.header("Dataset Overview")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Students", len(students_df))
        with col2:
            st.metric("Companies", len(companies_df))
        with col3:
            st.metric("Industries", companies_df['industry'].nunique())
        with col4:
            st.metric("Avg Ranking", f"{rankings_df['ranking'].mean():.2f}")
        
        # Ranking distribution
        st.header("Ranking Distribution")
        fig = px.histogram(rankings_df, x='ranking', nbins=10,
                          title="Distribution of All Rankings",
                          labels={'ranking': 'Ranking Score', 'count': 'Frequency'})
        st.plotly_chart(fig, use_container_width=True)
        
        # Company analysis
        st.header("Company Analysis")
        col1, col2 = st.columns(2)
        
        with col1:
            top_companies = companies_df.nlargest(min(15, len(companies_df)), 'it2_capacity')
            fig = px.bar(top_companies, x='company_name', y='it2_capacity',
                        title=f"Top {min(15, len(companies_df))} Companies by IT2 Capacity",
                        labels={'it2_capacity': 'IT2 Capacity', 'company_name': 'Company'},
                        color='industry',
                        color_discrete_sequence=px.colors.qualitative.Set3)
            fig.update_xaxes(tickangle=-45)
            st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            top_companies = companies_df.nlargest(min(15, len(companies_df)), 'it3_capacity')
            fig = px.bar(top_companies, x='company_name', y='it3_capacity',
                        title=f"Top {min(15, len(companies_df))} Companies by IT3 Capacity",
                        labels={'it3_capacity': 'IT3 Capacity', 'company_name': 'Company'},
                        color='industry',
                        color_discrete_sequence=px.colors.qualitative.Set3)
            fig.update_xaxes(tickangle=-45)
            st.plotly_chart(fig, use_container_width=True)
        
        # Total capacity comparison
        capacity_df = pd.DataFrame({
            'Capacity Type': ['IT2', 'IT3'],
            'Total Capacity': [companies_df['it2_capacity'].sum(), companies_df['it3_capacity'].sum()],
            'Students': [len(students_df), len(students_df)]
        })
        
        fig = go.Figure()
        fig.add_trace(go.Bar(name='Total Capacity', x=capacity_df['Capacity Type'], 
                             y=capacity_df['Total Capacity'], marker_color='lightblue'))
        fig.add_trace(go.Bar(name='Number of Students', x=capacity_df['Capacity Type'], 
                             y=capacity_df['Students'], marker_color='salmon'))
        fig.update_layout(title='Capacity vs Students Comparison',
                          xaxis_title='Iteration',
                          yaxis_title='Count',
                          barmode='group')
        st.plotly_chart(fig, use_container_width=True)
        
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

# PAGE 4: CLUSTERING
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
        
        # Create mapping of student IDs to names
        student_names_dict = students_df.set_index('student_id')['student_name'].to_dict()
        
        # Get the order of student IDs from dendrogram
        from scipy.cluster.hierarchy import dendrogram as scipy_dendrogram
        
        # First get the dendrogram with labels as student IDs
        dend = scipy_dendrogram(Z, labels=pivot_rankings.index.tolist(), no_plot=True)
        
        # Map the labels to student names
        student_id_labels = [int(x) for x in dend['ivl']]
        name_labels = [student_names_dict[sid] for sid in student_id_labels]
        
        # Create the plotly figure
        fig = go.Figure()
        
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
        
        # Update x-axis to show student names
        x_positions = dend['icoord']
        unique_x = sorted(set([x for coord in x_positions for x in coord if x % 10 == 5]))
        
        fig.update_layout(
            title=f"Hierarchical Clustering Dendrogram (Method: {method})",
            xaxis_title="Student Name",
            yaxis_title="Distance",
            height=600,
            xaxis=dict(
                tickmode='array',
                tickvals=unique_x,
                ticktext=name_labels,
                tickangle=-45
            )
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

# PAGE 5: OPTIMIZATION - UPDATED WITH ALL IMPROVEMENTS
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
        
        # PRE-OPTIMIZATION VALIDATION
        st.header("Pre-Optimization Checks")
        
        # Run feasibility checks
        is_feasible, feasibility_errors, feasibility_warnings = check_optimization_feasibility(students_df, companies_df)
        
        # Display validation results
        if feasibility_errors:
            st.error("🚫 **Critical Issues Detected:**")
            for err in feasibility_errors:
                st.write(err)
            st.error("**Cannot proceed with optimization. Please fix the issues above.**")
            st.stop()
        else:
            st.success("✅ Basic feasibility checks passed!")
        
        if feasibility_warnings:
            with st.expander("⚠️ View Warnings (Click to expand)", expanded=False):
                for warn in feasibility_warnings:
                    st.write(warn)
                st.info("💡 These warnings won't prevent optimization but may affect solution quality.")
        
        # Display industry capacity analysis
        st.subheader("Industry Capacity Analysis")
        industry_summary = companies_df.groupby('industry').agg({
            'it2_capacity': 'sum',
            'it3_capacity': 'sum',
            'company_id': 'count'
        }).rename(columns={'company_id': 'num_companies'})
        industry_summary['total_capacity'] = industry_summary['it2_capacity'] + industry_summary['it3_capacity']
        industry_summary = industry_summary.sort_values('total_capacity', ascending=False)
        
        st.dataframe(industry_summary, use_container_width=True)
        
        if st.button("🚀 Solve Optimization Problem", type="primary"):
            with st.spinner("Solving optimization problem..."):
                # Create the optimization problem
                prob = pulp.LpProblem("CoopPlacement", pulp.LpMaximize)
                
                students = students_df['student_id'].tolist()
                companies = companies_df['company_id'].tolist()
                
                # Decision variables
                x = {}  # IT2 assignments
                y = {}  # IT3 assignments
                
                for i in students:
                    for j in companies:
                        x[i, j] = pulp.LpVariable(f"x_{i}_{j}", cat='Binary')
                        y[i, j] = pulp.LpVariable(f"y_{i}_{j}", cat='Binary')
                
                # IMPROVED: Handle missing rankings by setting them to 0
                rankings_dict = rankings_df.set_index(['student_id', 'company_id'])['ranking'].to_dict()
                
                # Ensure all student-company pairs have a ranking (default to 0 if missing)
                complete_rankings_dict = {}
                for i in students:
                    for j in companies:
                        complete_rankings_dict[i, j] = rankings_dict.get((i, j), 0)
                
                # Objective function: Maximize total ranking
                prob += pulp.lpSum([complete_rankings_dict[i, j] * (x[i, j] + y[i, j]) 
                                   for i in students for j in companies])
                
                # Constraint 1: Every student does exactly one IT2 placement
                for i in students:
                    prob += pulp.lpSum([x[i, j] for j in companies]) == 1, f"IT2_assignment_student_{i}"
                
                # Constraint 2: Every student does exactly one IT3 placement
                for i in students:
                    prob += pulp.lpSum([y[i, j] for j in companies]) == 1, f"IT3_assignment_student_{i}"
                
                # Constraint 3: IT2 capacity constraints
                for j in companies:
                    capacity = companies_df[companies_df['company_id'] == j]['it2_capacity'].values[0]
                    prob += pulp.lpSum([x[i, j] for i in students]) <= capacity, f"IT2_capacity_company_{j}"
                
                # Constraint 4: IT3 capacity constraints
                for j in companies:
                    capacity = companies_df[companies_df['company_id'] == j]['it3_capacity'].values[0]
                    prob += pulp.lpSum([y[i, j] for i in students]) <= capacity, f"IT3_capacity_company_{j}"
                
                # Constraint 5: Students get different companies for IT2 and IT3
                for i in students:
                    for j in companies:
                        prob += x[i, j] + y[i, j] <= 1, f"Different_companies_student_{i}_company_{j}"
                
                # Constraint 6: Industry diversity - at most one placement per industry per student
                industry_companies = companies_df.groupby('industry')['company_id'].apply(list).to_dict()
                
                for i in students:
                    for industry, company_list in industry_companies.items():
                        prob += pulp.lpSum([x[i, j] + y[i, j] for j in company_list]) <= 1, \
                                f"Industry_diversity_student_{i}_industry_{industry}"
                
                # Solve the problem
                solver = pulp.PULP_CBC_CMD(msg=0)
                prob.solve(solver)
                
                # Display results
                status_text = pulp.LpStatus[prob.status]
                
                if prob.status == 1:  # Optimal solution found
                    st.success(f"✅ Optimization Status: {status_text}")
                else:
                    st.error(f"❌ Optimization Status: {status_text}")
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Objective Value", f"{pulp.value(prob.objective):.2f}")
                with col2:
                    max_possible = rankings_df['ranking'].max() * 2 * len(students)
                    st.metric("Max Possible", f"{max_possible}")
                with col3:
                    if pulp.value(prob.objective) and max_possible > 0:
                        efficiency = (pulp.value(prob.objective) / max_possible) * 100
                        st.metric("Efficiency", f"{efficiency:.1f}%")
                
                if prob.status == 1:
                    # Extract assignments
                    it2_assignments = []
                    it3_assignments = []
                    
                    for i in students:
                        for j in companies:
                            if pulp.value(x[i, j]) == 1:
                                student_name = students_df[students_df['student_id'] == i]['student_name'].values[0]
                                company_name = companies_df[companies_df['company_id'] == j]['company_name'].values[0]
                                industry = companies_df[companies_df['company_id'] == j]['industry'].values[0]
                                ranking = complete_rankings_dict[i, j]
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
                                ranking = complete_rankings_dict[i, j]
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
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("Average IT2 Ranking", f"{it2_df['Ranking'].mean():.2f}")
                        with col2:
                            st.metric("Min IT2 Ranking", f"{it2_df['Ranking'].min():.0f}")
                        with col3:
                            st.metric("Max IT2 Ranking", f"{it2_df['Ranking'].max():.0f}")
                    
                    with tab2:
                        st.subheader("IT3 Placements")
                        st.dataframe(it3_df.sort_values('Student'), hide_index=True, height=500)
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("Average IT3 Ranking", f"{it3_df['Ranking'].mean():.2f}")
                        with col2:
                            st.metric("Min IT3 Ranking", f"{it3_df['Ranking'].min():.0f}")
                        with col3:
                            st.metric("Max IT3 Ranking", f"{it3_df['Ranking'].max():.0f}")
                    
                    with tab3:
                        combined = it2_df.merge(it3_df, on='Student', suffixes=('_IT2', '_IT3'))
                        st.dataframe(combined, hide_index=True, height=500)
                        
                        # Check diversity constraint satisfaction
                        st.subheader("Diversity Verification")
                        diversity_check = []
                        for _, row in combined.iterrows():
                            same_industry = row['Industry_IT2'] == row['Industry_IT3']
                            diversity_check.append({
                                'Student': row['Student'],
                                'IT2_Industry': row['Industry_IT2'],
                                'IT3_Industry': row['Industry_IT3'],
                                'Diverse': '✅' if not same_industry else '❌'
                            })
                        
                        diversity_df = pd.DataFrame(diversity_check)
                        all_diverse = all(row['Diverse'] == '✅' for row in diversity_check)
                        
                        if all_diverse:
                            st.success("✅ All students have diverse industry placements!")
                        else:
                            st.warning("⚠️ Some students have placements in the same industry")
                            st.dataframe(diversity_df[diversity_df['Diverse'] == '❌'])
                    
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
                    
                    # Additional analytics
                    st.subheader("Student Satisfaction Analysis")
                    combined['Total_Ranking'] = combined['Ranking_IT2'] + combined['Ranking_IT3']
                    combined['Avg_Ranking'] = combined['Total_Ranking'] / 2
                    
                    fig = px.bar(combined.sort_values('Total_Ranking', ascending=False).head(20),
                               x='Student', y='Total_Ranking',
                               title="Top 20 Students by Total Ranking",
                               color='Avg_Ranking',
                               color_continuous_scale='RdYlGn')
                    fig.update_xaxes(tickangle=-45)
                    st.plotly_chart(fig, use_container_width=True)
                    
                    st.header("Export Results")
                    
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        it2_df.to_excel(writer, sheet_name='IT2_Placements', index=False)
                        it3_df.to_excel(writer, sheet_name='IT3_Placements', index=False)
                        combined.to_excel(writer, sheet_name='Combined', index=False)
                        
                        # Add summary sheet
                        summary_data = {
                            'Metric': [
                                'Total Students',
                                'Total Companies',
                                'Objective Value',
                                'Max Possible',
                                'Efficiency (%)',
                                'Avg IT2 Ranking',
                                'Avg IT3 Ranking',
                                'Overall Avg Ranking'
                            ],
                            'Value': [
                                len(students),
                                len(companies),
                                f"{pulp.value(prob.objective):.2f}",
                                max_possible,
                                f"{efficiency:.1f}" if 'efficiency' in locals() else 'N/A',
                                f"{it2_df['Ranking'].mean():.2f}",
                                f"{it3_df['Ranking'].mean():.2f}",
                                f"{pd.concat([it2_df['Ranking'], it3_df['Ranking']]).mean():.2f}"
                            ]
                        }
                        summary_df = pd.DataFrame(summary_data)
                        summary_df.to_excel(writer, sheet_name='Summary', index=False)
                    
                    st.download_button(
                        label="📥 Download Results (Excel)",
                        data=output.getvalue(),
                        file_name="placement_results.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("❌ Optimization failed to find a solution.")
                    
                    st.subheader("Possible Reasons:")
                    reasons = []
                    
                    # Check capacity issues
                    if companies_df['it2_capacity'].sum() < n_students:
                        reasons.append("• Total IT2 capacity is less than number of students")
                    if companies_df['it3_capacity'].sum() < n_students:
                        reasons.append("• Total IT3 capacity is less than number of students")
                    
                    # Check industry diversity feasibility
                    industry_capacities = companies_df.groupby('industry')[['it2_capacity', 'it3_capacity']].sum()
                    for industry, row in industry_capacities.iterrows():
                        total_cap = row['it2_capacity'] + row['it3_capacity']
                        if total_cap < n_students:
                            reasons.append(f"• Industry '{industry}' has insufficient total capacity ({int(total_cap)} < {n_students})")
                    
                    if not reasons:
                        reasons.append("• Unknown issue - please check your data")
                    
                    for reason in reasons:
                        st.write(reason)
                    
                    st.info("💡 Try adjusting company capacities or industry distribution to make the problem feasible.")

# Footer
st.sidebar.markdown("---")
st.sidebar.info(f"""
**Co-op Placement Optimizer**  
Version 4.0 - Enhanced Validation
Built with Streamlit & PuLP

Current Data:
- Students: {len(st.session_state.students_df) if st.session_state.data_loaded else 0}
- Companies: {len(st.session_state.companies_df) if st.session_state.data_loaded else 0}
- Industries: {st.session_state.companies_df['industry'].nunique() if st.session_state.data_loaded and len(st.session_state.companies_df) > 0 else 0}

Features:
✅ Data validation
✅ Feasibility checking
✅ Missing ranking handling
✅ Industry capacity analysis
✅ Enhanced error messages
""")
