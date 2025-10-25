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
                        ["Data Setup", "Manual Data Editor", "Rankings Editor", "Exploratory Analysis", "Clustering", "Optimization"],
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
    """Generate default rankings for all student-company pairs"""
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
    """Load data from Excel file"""
    students_df = pd.read_excel(file, sheet_name='Students')
    companies_df = pd.read_excel(file, sheet_name='Companies')
    rankings_df = pd.read_excel(file, sheet_name='Rankings')
    return students_df, companies_df, rankings_df

# =============================================================================
# PAGE: DATA SETUP
# =============================================================================
if page == "Data Setup":
    st.title("📊 Data Setup")
    
    st.write("""
    Choose how to load your data:
    - **Upload Excel File**: Use a pre-filled template
    - **Generate Synthetic Data**: For testing and demonstration
    - **Start Fresh**: Create data manually in the editor
    """)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.subheader("📁 Upload Excel")
        uploaded_file = st.file_uploader("Upload Excel file", type=['xlsx'], key='upload')
        if uploaded_file and st.button("Load Excel Data"):
            try:
                students, companies, rankings = load_excel_data(uploaded_file)
                st.session_state.students_df = students
                st.session_state.companies_df = companies
                st.session_state.rankings_df = rankings
                st.session_state.data_loaded = True
                st.success("✅ Data loaded successfully!")
                st.rerun()
            except Exception as e:
                st.error(f"Error loading file: {e}")
    
    with col2:
        st.subheader("🎲 Generate Synthetic Data")
        n_students_synth = st.number_input("Number of students", 5, 100, 15, key='n_students_synth')
        n_companies_synth = st.number_input("Number of companies", 5, 100, 15, key='n_companies_synth')
        if st.button("Generate Data"):
            students, companies, rankings = generate_synthetic_data(n_students_synth, n_companies_synth)
            st.session_state.students_df = students
            st.session_state.companies_df = companies
            st.session_state.rankings_df = rankings
            st.session_state.data_loaded = True
            st.success("✅ Synthetic data generated!")
            st.rerun()
    
    with col3:
        st.subheader("✍️ Start Fresh")
        if st.button("Create Empty Dataset"):
            students, companies, rankings = initialize_empty_data()
            st.session_state.students_df = students
            st.session_state.companies_df = companies
            st.session_state.rankings_df = rankings
            st.session_state.data_loaded = True
            st.success("✅ Empty dataset created!")
            st.rerun()
    
    st.markdown("---")
    st.subheader("📥 Download Excel Template")
    
    col1, col2 = st.columns(2)
    with col1:
        template_students = st.number_input("Students in template", 5, 100, 20, key='template_students')
    with col2:
        template_companies = st.number_input("Companies in template", 5, 100, 20, key='template_companies')
    
    template_buffer = create_excel_template(template_students, template_companies)
    st.download_button(
        label="📥 Download Template",
        data=template_buffer,
        file_name="coop_placement_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    if st.session_state.data_loaded:
        st.markdown("---")
        st.subheader("📊 Current Data Summary")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Students", len(st.session_state.students_df))
        with col2:
            st.metric("Companies", len(st.session_state.companies_df))
        with col3:
            st.metric("Rankings", len(st.session_state.rankings_df))

# =============================================================================
# PAGE: MANUAL DATA EDITOR (Improved)
# =============================================================================
elif page == "Manual Data Editor":
    st.title("✍️ Manual Data Editor")
    
    if not st.session_state.data_loaded:
        st.warning("⚠️ Please load or create data first in 'Data Setup'")
    else:
        st.write("Edit students and companies information. Student names can be any text (names, numbers, or both).")
        
        tab1, tab2 = st.tabs(["👥 Students", "🏢 Companies"])
        
        with tab1:
            st.subheader("Student Information")
            st.info("💡 Student names can be any format: 'John Smith', 'Student 1', '2024-S-001', etc.")
            
            # Add new student
            with st.expander("➕ Add New Student", expanded=False):
                col1, col2 = st.columns(2)
                with col1:
                    new_student_name = st.text_input("Student Name (any format)", 
                                                     placeholder="e.g., Sarah Johnson, Student01, or 2024-001")
                with col2:
                    new_student_gpa = st.number_input("GPA", 0.0, 4.0, 3.5, 0.01)
                
                if st.button("Add Student"):
                    if new_student_name.strip():
                        new_id = st.session_state.students_df['student_id'].max() + 1 if len(st.session_state.students_df) > 0 else 1
                        new_student = pd.DataFrame({
                            'student_id': [new_id],
                            'student_name': [new_student_name.strip()],
                            'gpa': [new_student_gpa]
                        })
                        st.session_state.students_df = pd.concat([st.session_state.students_df, new_student], ignore_index=True)
                        
                        # Generate rankings for new student
                        companies = st.session_state.companies_df['company_id'].tolist()
                        new_rankings = pd.DataFrame([
                            {'student_id': new_id, 'company_id': comp_id, 'ranking': 5}
                            for comp_id in companies
                        ])
                        st.session_state.rankings_df = pd.concat([st.session_state.rankings_df, new_rankings], ignore_index=True)
                        
                        st.success(f"✅ Added student: {new_student_name}")
                        st.rerun()
                    else:
                        st.error("Please enter a student name")
            
            # Edit existing students
            st.markdown("#### Current Students")
            edited_students = st.data_editor(
                st.session_state.students_df,
                use_container_width=True,
                num_rows="dynamic",
                column_config={
                    "student_id": st.column_config.NumberColumn("ID", disabled=True),
                    "student_name": st.column_config.TextColumn("Name", help="Any format accepted"),
                    "gpa": st.column_config.NumberColumn("GPA", min_value=0.0, max_value=4.0, format="%.2f")
                },
                hide_index=True
            )
            
            if st.button("💾 Save Student Changes"):
                st.session_state.students_df = edited_students
                # Regenerate rankings if needed
                if len(st.session_state.companies_df) > 0:
                    st.session_state.rankings_df = regenerate_rankings(
                        st.session_state.students_df, 
                        st.session_state.companies_df
                    )
                st.success("✅ Student data saved!")
                st.rerun()
        
        with tab2:
            st.subheader("Company Information")
            
            # Add new company
            with st.expander("➕ Add New Company", expanded=False):
                col1, col2 = st.columns(2)
                with col1:
                    new_company_name = st.text_input("Company Name")
                    new_industry = st.selectbox("Industry", 
                                               ["General Insurance", "Consultancy", "Life Insurance", "Care/Disability"])
                with col2:
                    new_it2_cap = st.number_input("IT2 Capacity", 1, 100, 1)
                    new_it3_cap = st.number_input("IT3 Capacity", 1, 100, 1)
                
                if st.button("Add Company"):
                    if new_company_name.strip():
                        new_id = st.session_state.companies_df['company_id'].max() + 1 if len(st.session_state.companies_df) > 0 else 1
                        new_company = pd.DataFrame({
                            'company_id': [new_id],
                            'company_name': [new_company_name],
                            'industry': [new_industry],
                            'it2_capacity': [new_it2_cap],
                            'it3_capacity': [new_it3_cap]
                        })
                        st.session_state.companies_df = pd.concat([st.session_state.companies_df, new_company], ignore_index=True)
                        
                        # Generate rankings for new company
                        students = st.session_state.students_df['student_id'].tolist()
                        new_rankings = pd.DataFrame([
                            {'student_id': stud_id, 'company_id': new_id, 'ranking': 5}
                            for stud_id in students
                        ])
                        st.session_state.rankings_df = pd.concat([st.session_state.rankings_df, new_rankings], ignore_index=True)
                        
                        st.success(f"✅ Added company: {new_company_name}")
                        st.rerun()
                    else:
                        st.error("Please enter a company name")
            
            # Edit existing companies
            st.markdown("#### Current Companies")
            edited_companies = st.data_editor(
                st.session_state.companies_df,
                use_container_width=True,
                num_rows="dynamic",
                column_config={
                    "company_id": st.column_config.NumberColumn("ID", disabled=True),
                    "company_name": st.column_config.TextColumn("Company Name"),
                    "industry": st.column_config.SelectboxColumn("Industry",
                        options=["General Insurance", "Consultancy", "Life Insurance", "Care/Disability"]),
                    "it2_capacity": st.column_config.NumberColumn("IT2 Capacity", min_value=0),
                    "it3_capacity": st.column_config.NumberColumn("IT3 Capacity", min_value=0)
                },
                hide_index=True
            )
            
            if st.button("💾 Save Company Changes"):
                st.session_state.companies_df = edited_companies
                # Regenerate rankings if needed
                if len(st.session_state.students_df) > 0:
                    st.session_state.rankings_df = regenerate_rankings(
                        st.session_state.students_df, 
                        st.session_state.companies_df
                    )
                st.success("✅ Company data saved!")
                st.rerun()

# =============================================================================
# PAGE: RANKINGS EDITOR (New!)
# =============================================================================
elif page == "Rankings Editor":
    st.title("📊 Rankings Editor")
    
    if not st.session_state.data_loaded:
        st.warning("⚠️ Please load or create data first in 'Data Setup'")
    elif len(st.session_state.students_df) == 0 or len(st.session_state.companies_df) == 0:
        st.warning("⚠️ Please add students and companies first in 'Manual Data Editor'")
    else:
        st.write("Edit student rankings for companies. **Higher ranking = stronger preference** (scale: 1-10)")
        
        # Create pivot table for easier editing
        students_df = st.session_state.students_df
        companies_df = st.session_state.companies_df
        rankings_df = st.session_state.rankings_df
        
        # Merge to get names
        rankings_with_names = rankings_df.merge(
            students_df[['student_id', 'student_name']], on='student_id'
        ).merge(
            companies_df[['company_id', 'company_name']], on='company_id'
        )
        
        # Method selection
        edit_method = st.radio("Choose editing method:", 
                              ["Matrix View (All at once)", "Student-by-Student View"])
        
        if edit_method == "Matrix View (All at once)":
            st.info("💡 Edit the entire ranking matrix. Each cell is a student's ranking for a company.")
            
            # Create pivot table
            rankings_pivot = rankings_with_names.pivot(
                index='student_name',
                columns='company_name',
                values='ranking'
            )
            
            # Display editable dataframe
            edited_rankings_pivot = st.data_editor(
                rankings_pivot,
                use_container_width=True,
                column_config={col: st.column_config.NumberColumn(
                    col,
                    min_value=1,
                    max_value=10,
                    format="%d"
                ) for col in rankings_pivot.columns}
            )
            
            if st.button("💾 Save All Rankings", type="primary"):
                # Convert back to long format
                new_rankings_list = []
                for student_name in edited_rankings_pivot.index:
                    student_id = students_df[students_df['student_name'] == student_name]['student_id'].values[0]
                    for company_name in edited_rankings_pivot.columns:
                        company_id = companies_df[companies_df['company_name'] == company_name]['company_id'].values[0]
                        ranking = edited_rankings_pivot.loc[student_name, company_name]
                        new_rankings_list.append({
                            'student_id': student_id,
                            'company_id': company_id,
                            'ranking': int(ranking)
                        })
                
                st.session_state.rankings_df = pd.DataFrame(new_rankings_list)
                st.success("✅ All rankings saved successfully!")
                st.rerun()
        
        else:  # Student-by-Student View
            st.info("💡 Edit rankings one student at a time for more focused input.")
            
            # Select student
            student_names = students_df['student_name'].tolist()
            selected_student_name = st.selectbox("Select Student", student_names)
            selected_student_id = students_df[students_df['student_name'] == selected_student_name]['student_id'].values[0]
            
            # Get rankings for this student
            student_rankings = rankings_with_names[
                rankings_with_names['student_id'] == selected_student_id
            ][['company_name', 'industry', 'ranking']].copy()
            
            st.markdown(f"#### Rankings for: **{selected_student_name}**")
            
            # Edit rankings for this student
            edited_student_rankings = st.data_editor(
                student_rankings,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "company_name": st.column_config.TextColumn("Company", disabled=True),
                    "industry": st.column_config.TextColumn("Industry", disabled=True),
                    "ranking": st.column_config.NumberColumn(
                        "Ranking",
                        min_value=1,
                        max_value=10,
                        help="1 = lowest preference, 10 = highest preference",
                        format="%d"
                    )
                }
            )
            
            col1, col2 = st.columns([3, 1])
            with col1:
                if st.button("💾 Save Rankings for This Student", type="primary"):
                    # Update rankings for this student
                    for idx, row in edited_student_rankings.iterrows():
                        company_id = companies_df[companies_df['company_name'] == row['company_name']]['company_id'].values[0]
                        mask = (rankings_df['student_id'] == selected_student_id) & (rankings_df['company_id'] == company_id)
                        st.session_state.rankings_df.loc[mask, 'ranking'] = int(row['ranking'])
                    
                    st.success(f"✅ Rankings saved for {selected_student_name}!")
                    st.rerun()
            
            with col2:
                if st.button("🔄 Reset to 5"):
                    edited_student_rankings['ranking'] = 5
                    st.rerun()
            
            # Quick stats
            st.markdown("---")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Average Ranking", f"{edited_student_rankings['ranking'].mean():.1f}")
            with col2:
                st.metric("Highest Ranking", f"{edited_student_rankings['ranking'].max()}")
            with col3:
                st.metric("Lowest Ranking", f"{edited_student_rankings['ranking'].min()}")

# =============================================================================
# PAGE: EXPLORATORY ANALYSIS
# =============================================================================
elif page == "Exploratory Analysis":
    st.title("📈 Exploratory Analysis")
    
    if not st.session_state.data_loaded:
        st.warning("⚠️ Please load or create data first in 'Data Setup'")
    else:
        students_df = st.session_state.students_df
        companies_df = st.session_state.companies_df
        rankings_df = st.session_state.rankings_df
        
        # Merge dataframes
        full_data = rankings_df.merge(students_df, on='student_id').merge(companies_df, on='company_id')
        
        tab1, tab2, tab3 = st.tabs(["Overview", "Student Analysis", "Company Analysis"])
        
        with tab1:
            st.subheader("Dataset Overview")
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total Students", len(students_df))
            with col2:
                st.metric("Total Companies", len(companies_df))
            with col3:
                st.metric("Average Ranking", f"{rankings_df['ranking'].mean():.2f}")
            with col4:
                total_capacity = companies_df['it2_capacity'].sum() + companies_df['it3_capacity'].sum()
                needed = len(students_df) * 2
                st.metric("Capacity Ratio", f"{total_capacity / needed:.2f}" if needed > 0 else "N/A")
            
            col1, col2 = st.columns(2)
            with col1:
                fig = px.histogram(rankings_df, x='ranking', nbins=10,
                                 title="Distribution of All Rankings")
                st.plotly_chart(fig, use_container_width=True)
            
            with col2:
                industry_counts = companies_df['industry'].value_counts()
                fig = px.pie(values=industry_counts.values, names=industry_counts.index,
                           title="Companies by Industry")
                st.plotly_chart(fig, use_container_width=True)
        
        with tab2:
            st.subheader("Student Preference Analysis")
            
            student_avg = full_data.groupby('student_name')['ranking'].agg(['mean', 'std', 'min', 'max']).reset_index()
            student_avg.columns = ['Student', 'Avg Ranking', 'Std Dev', 'Min', 'Max']
            
            fig = px.bar(student_avg.sort_values('Avg Ranking', ascending=False), 
                        x='Student', y='Avg Ranking',
                        title="Average Ranking by Student")
            st.plotly_chart(fig, use_container_width=True)
            
            st.dataframe(student_avg, hide_index=True, use_container_width=True)
        
        with tab3:
            st.subheader("Company Popularity Analysis")
            
            company_avg = full_data.groupby('company_name')['ranking'].agg(['mean', 'std', 'count']).reset_index()
            company_avg.columns = ['Company', 'Avg Ranking', 'Std Dev', 'Count']
            
            fig = px.bar(company_avg.sort_values('Avg Ranking', ascending=False), 
                        x='Company', y='Avg Ranking',
                        title="Average Ranking Received by Company")
            st.plotly_chart(fig, use_container_width=True)
            
            st.dataframe(company_avg, hide_index=True, use_container_width=True)

# =============================================================================
# PAGE: CLUSTERING
# =============================================================================
elif page == "Clustering":
    st.title("🔍 Clustering Analysis")
    
    if not st.session_state.data_loaded:
        st.warning("⚠️ Please load or create data first in 'Data Setup'")
    else:
        st.write("Analyze similarity patterns among students and companies based on rankings.")
        
        students_df = st.session_state.students_df
        companies_df = st.session_state.companies_df
        rankings_df = st.session_state.rankings_df
        
        cluster_type = st.radio("Cluster by:", ["Students", "Companies"])
        
        if cluster_type == "Students":
            # Pivot: rows=students, cols=companies
            pivot_df = rankings_df.pivot(index='student_id', columns='company_id', values='ranking')
            pivot_df = pivot_df.fillna(pivot_df.mean())
            
            # Merge student names
            pivot_df = pivot_df.merge(students_df[['student_id', 'student_name']], 
                                     left_index=True, right_on='student_id', how='left')
            pivot_df.set_index('student_name', inplace=True)
            pivot_df.drop('student_id', axis=1, inplace=True)
            
            st.subheader("Student Clustering")
            n_clusters = st.slider("Number of clusters", 2, min(10, len(students_df)), 3)
            
            if st.button("Run Clustering"):
                linkage_matrix = linkage(pivot_df.values, method='ward')
                clusters = fcluster(linkage_matrix, n_clusters, criterion='maxclust')
                
                fig = go.Figure(data=go.Heatmap(
                    z=pivot_df.values,
                    x=[f"C{i}" for i in pivot_df.columns],
                    y=pivot_df.index,
                    colorscale='RdYlGn'
                ))
                fig.update_layout(title="Student-Company Rankings Heatmap")
                st.plotly_chart(fig, use_container_width=True)
                
                cluster_df = pd.DataFrame({
                    'Student': pivot_df.index,
                    'Cluster': clusters
                })
                st.dataframe(cluster_df, hide_index=True)
        
        else:  # Companies
            # Pivot: rows=companies, cols=students
            pivot_df = rankings_df.pivot(index='company_id', columns='student_id', values='ranking')
            pivot_df = pivot_df.fillna(pivot_df.mean())
            
            # Merge company names
            pivot_df = pivot_df.merge(companies_df[['company_id', 'company_name']], 
                                     left_index=True, right_on='company_id', how='left')
            pivot_df.set_index('company_name', inplace=True)
            pivot_df.drop('company_id', axis=1, inplace=True)
            
            st.subheader("Company Clustering")
            n_clusters = st.slider("Number of clusters", 2, min(10, len(companies_df)), 3)
            
            if st.button("Run Clustering"):
                linkage_matrix = linkage(pivot_df.values, method='ward')
                clusters = fcluster(linkage_matrix, n_clusters, criterion='maxclust')
                
                fig = go.Figure(data=go.Heatmap(
                    z=pivot_df.values,
                    x=[f"S{i}" for i in pivot_df.columns],
                    y=pivot_df.index,
                    colorscale='RdYlGn'
                ))
                fig.update_layout(title="Company-Student Rankings Heatmap")
                st.plotly_chart(fig, use_container_width=True)
                
                cluster_df = pd.DataFrame({
                    'Company': pivot_df.index,
                    'Cluster': clusters
                })
                st.dataframe(cluster_df, hide_index=True)

# =============================================================================
# PAGE: OPTIMIZATION
# =============================================================================
elif page == "Optimization":
    st.title("🎯 Optimization")
    
    if not st.session_state.data_loaded:
        st.warning("⚠️ Please load or create data first in 'Data Setup'")
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
Version 4.0 - Enhanced Editor
Built with Streamlit & PuLP

Current Data:
- Students: {len(st.session_state.students_df) if st.session_state.data_loaded else 0}
- Companies: {len(st.session_state.companies_df) if st.session_state.data_loaded else 0}
""")
