import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import pulp
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ============================================================================
# PAGE CONFIGURATION
# ============================================================================
st.set_page_config(
    page_title="Co-op Placement Optimizer",
    layout="wide",
    page_icon="🎓",
    initial_sidebar_state="collapsed"
)

# Custom CSS for better styling
st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #666;
        text-align: center;
        margin-bottom: 2rem;
    }
    .step-card {
        background-color: #f0f2f6;
        padding: 1.5rem;
        border-radius: 10px;
        margin: 1rem 0;
    }
    .success-box {
        background-color: #d4edda;
        border-left: 5px solid #28a745;
        padding: 1rem;
        border-radius: 5px;
        margin: 1rem 0;
    }
    .info-box {
        background-color: #d1ecf1;
        border-left: 5px solid #17a2b8;
        padding: 1rem;
        border-radius: 5px;
        margin: 1rem 0;
    }
    .warning-box {
        background-color: #fff3cd;
        border-left: 5px solid #ffc107;
        padding: 1rem;
        border-radius: 5px;
        margin: 1rem 0;
    }
    </style>
    """, unsafe_allow_html=True)

# ============================================================================
# SESSION STATE INITIALIZATION
# ============================================================================
if 'step' not in st.session_state:
    st.session_state.step = 1
if 'data_loaded' not in st.session_state:
    st.session_state.data_loaded = False
if 'students_df' not in st.session_state:
    st.session_state.students_df = None
if 'companies_df' not in st.session_state:
    st.session_state.companies_df = None
if 'rankings_df' not in st.session_state:
    st.session_state.rankings_df = None
if 'wide_format_df' not in st.session_state:
    st.session_state.wide_format_df = None
if 'optimization_done' not in st.session_state:
    st.session_state.optimization_done = False

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def create_template(n_students=11, n_companies=11):
    """Create simple Excel template"""
    
    # Default names
    student_names = ['Aiden', 'Angie', 'George', 'James', 'Josh', 
                    'Kalea', 'Kenzie', 'Prapann', 'Sofia', 'Tony', 'Vihaan']
    
    if n_students > len(student_names):
        student_names += [f'Student_{i}' for i in range(len(student_names)+1, n_students+1)]
    else:
        student_names = student_names[:n_students]
    
    companies_data = [
        ('Consultancy', 'EY', 1, 1),
        ('Consultancy', 'Finity', 1, 1),
        ('Consultancy', 'PwC', 1, 1),
        ('General Insurance', 'Allianz', 1, 1),
        ('General Insurance', 'IAG', 1, 1),
        ('General Insurance', 'Suncorp', 1, 1),
        ('Life Insurance', 'AMP', 1, 1),
        ('Life Insurance', 'TAL', 1, 1),
        ('Government', 'APRA', 1, 1),
        ('Government', 'icare', 1, 1),
        ('Government', 'NDIA', 1, 1),
    ]
    
    # Create workbook
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Placement Data"
    
    # Headers
    headers = ['Industry', 'Company', 'IT2 Capacity', 'IT3 Capacity'] + student_names
    
    # Styling
    header_fill = PatternFill(start_color='1f77b4', end_color='1f77b4', fill_type='solid')
    header_font = Font(bold=True, color='FFFFFF', size=12)
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Write headers
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(1, col_idx, header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border
        ws.row_dimensions[1].height = 30
    
    # Fill data
    for row_idx, (industry, company, it2, it3) in enumerate(companies_data[:n_companies], 2):
        ws.cell(row_idx, 1, industry).border = border
        ws.cell(row_idx, 2, company).border = border
        ws.cell(row_idx, 3, it2).border = border
        ws.cell(row_idx, 4, it3).border = border
        
        for col_idx in range(5, 5 + len(student_names)):
            cell = ws.cell(row_idx, col_idx)
            cell.border = border
            cell.alignment = Alignment(horizontal='center')
    
    # Column widths
    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 12
    ws.column_dimensions['D'].width = 12
    for col_idx in range(5, 5 + len(student_names)):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = 11
    
    # Instructions
    inst_row = n_companies + 3
    ws.cell(inst_row, 1, "HOW TO FILL THIS TEMPLATE:").font = Font(bold=True, size=14)
    ws.cell(inst_row+1, 1, "1️⃣  Industry: Type the industry category (e.g., Consultancy, Insurance)")
    ws.cell(inst_row+2, 1, "2️⃣  Company: Enter company names")
    ws.cell(inst_row+3, 1, "3️⃣  IT2/IT3 Capacity: How many students can the company take? (Use 0 if not available)")
    ws.cell(inst_row+4, 1, "4️⃣  Student columns: Each student ranks each company from 1-10")
    ws.cell(inst_row+5, 1, "     • 10 = Really want to work here!")
    ws.cell(inst_row+6, 1, "     • 1 = Really don't want to work here")
    ws.cell(inst_row+7, 1, "     • Leave empty = No preference (treated as 0)")
    ws.cell(inst_row+8, 1, "5️⃣  Save and upload to the app!")
    
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

def parse_excel(file):
    """Parse uploaded Excel file"""
    try:
        df = pd.read_excel(file, sheet_name=0)
        
        # Check columns
        required = ['Industry', 'Company', 'IT2 Capacity', 'IT3 Capacity']
        if not all(col in df.columns for col in required):
            return None, None, None, None, f"❌ Missing columns. Need: {', '.join(required)}"
        
        df = df.dropna(subset=['Company'])
        student_cols = [col for col in df.columns if col not in required]
        
        if len(student_cols) == 0:
            return None, None, None, None, "❌ No student columns found"
        
        # Create dataframes
        students_df = pd.DataFrame({
            'student_id': range(1, len(student_cols) + 1),
            'student_name': student_cols
        })
        
        companies_df = pd.DataFrame({
            'company_id': range(1, len(df) + 1),
            'company_name': df['Company'].values,
            'industry': df['Industry'].values,
            'it2_capacity': df['IT2 Capacity'].fillna(0).astype(int).values,
            'it3_capacity': df['IT3 Capacity'].fillna(0).astype(int).values
        })
        
        # Rankings
        rankings_data = []
        for company_idx, company_id in enumerate(companies_df['company_id']):
            for student_idx, student_col in enumerate(student_cols):
                student_id = student_idx + 1
                ranking = df[student_col].iloc[company_idx]
                ranking = int(ranking) if pd.notna(ranking) else 0
                rankings_data.append({
                    'student_id': student_id,
                    'company_id': company_id,
                    'ranking': ranking
                })
        
        rankings_df = pd.DataFrame(rankings_data)
        wide_format_df = df.copy()
        
        return students_df, companies_df, rankings_df, wide_format_df, None
        
    except Exception as e:
        return None, None, None, None, f"❌ Error: {str(e)}"

def check_feasibility(students_df, companies_df):
    """Check if problem is solvable"""
    errors = []
    warnings = []
    n_students = len(students_df)
    
    total_it2 = companies_df['it2_capacity'].sum()
    total_it3 = companies_df['it3_capacity'].sum()
    
    if total_it2 < n_students:
        errors.append(f"Not enough IT2 spots! Need {n_students}, have {total_it2}")
    
    if total_it3 < n_students:
        errors.append(f"Not enough IT3 spots! Need {n_students}, have {total_it3}")
    
    # Check industries
    industry_caps = companies_df.groupby('industry')[['it2_capacity', 'it3_capacity']].sum()
    for industry, row in industry_caps.iterrows():
        total = row['it2_capacity'] + row['it3_capacity']
        if total < n_students:
            errors.append(f"Industry '{industry}' doesn't have enough total spots ({int(total)} < {n_students})")
    
    if total_it2 < n_students * 1.5:
        warnings.append(f"⚠️ IT2 capacity is tight ({total_it2} spots for {n_students} students)")
    if total_it3 < n_students * 1.5:
        warnings.append(f"⚠️ IT3 capacity is tight ({total_it3} spots for {n_students} students)")
    
    return len(errors) == 0, errors, warnings

def solve_optimization(students_df, companies_df, rankings_df):
    """Solve the placement problem"""
    
    prob = pulp.LpProblem("CoopPlacement", pulp.LpMaximize)
    
    students = students_df['student_id'].tolist()
    companies = companies_df['company_id'].tolist()
    
    # Variables
    x = {(i, j): pulp.LpVariable(f"x_{i}_{j}", cat='Binary') 
         for i in students for j in companies}
    y = {(i, j): pulp.LpVariable(f"y_{i}_{j}", cat='Binary') 
         for i in students for j in companies}
    
    # Rankings
    rankings_dict = rankings_df.set_index(['student_id', 'company_id'])['ranking'].to_dict()
    rankings_dict = {(i, j): rankings_dict.get((i, j), 0) for i in students for j in companies}
    
    # Objective: Maximize satisfaction
    prob += pulp.lpSum([rankings_dict[i, j] * (x[i, j] + y[i, j]) 
                       for i in students for j in companies])
    
    # Constraints
    for i in students:
        prob += pulp.lpSum([x[i, j] for j in companies]) == 1
        prob += pulp.lpSum([y[i, j] for j in companies]) == 1
    
    for j in companies:
        cap_it2 = companies_df[companies_df['company_id'] == j]['it2_capacity'].values[0]
        cap_it3 = companies_df[companies_df['company_id'] == j]['it3_capacity'].values[0]
        prob += pulp.lpSum([x[i, j] for i in students]) <= cap_it2
        prob += pulp.lpSum([y[i, j] for i in students]) <= cap_it3
    
    for i in students:
        for j in companies:
            prob += x[i, j] + y[i, j] <= 1
    
    # Industry diversity
    industry_companies = companies_df.groupby('industry')['company_id'].apply(list).to_dict()
    for i in students:
        for industry, company_list in industry_companies.items():
            prob += pulp.lpSum([x[i, j] + y[i, j] for j in company_list]) <= 1
    
    # Solve
    solver = pulp.PULP_CBC_CMD(msg=0)
    prob.solve(solver)
    
    return prob, x, y, rankings_dict

# ============================================================================
# MAIN APP
# ============================================================================

# Header
st.markdown('<p class="main-header">🎓 Co-op Placement Optimizer</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Find the best IT2 and IT3 placements for your students in 3 easy steps!</p>', unsafe_allow_html=True)

# Progress bar
progress_percent = (st.session_state.step - 1) / 2
st.progress(progress_percent)

# Step indicator
col1, col2, col3 = st.columns(3)
with col1:
    if st.session_state.step >= 1:
        st.markdown("### ✅ Step 1: Upload Data")
    else:
        st.markdown("### 1️⃣ Step 1: Upload Data")
with col2:
    if st.session_state.step >= 2:
        st.markdown("### ✅ Step 2: Review & Edit")
    else:
        st.markdown("### 2️⃣ Step 2: Review & Edit")
with col3:
    if st.session_state.step >= 3:
        st.markdown("### ✅ Step 3: Get Results")
    else:
        st.markdown("### 3️⃣ Step 3: Get Results")

st.markdown("---")

# ============================================================================
# STEP 1: UPLOAD DATA
# ============================================================================

if st.session_state.step == 1:
    st.markdown("## 📤 Step 1: Get Your Data Ready")
    
    tab1, tab2 = st.tabs(["📥 Upload Existing File", "📝 Create New Template"])
    
    with tab1:
        st.markdown("### Upload Your Excel File")
        st.info("💡 **Your file should have:** Industry, Company, IT2 Capacity, IT3 Capacity columns, plus one column per student with their rankings (1-10)")
        
        uploaded_file = st.file_uploader("Choose Excel file", type=['xlsx'], key='uploader')
        
        if uploaded_file is not None:
            with st.spinner("📖 Reading your file..."):
                students_df, companies_df, rankings_df, wide_format_df, error = parse_excel(uploaded_file)
                
                if error:
                    st.error(error)
                else:
                    st.session_state.students_df = students_df
                    st.session_state.companies_df = companies_df
                    st.session_state.rankings_df = rankings_df
                    st.session_state.wide_format_df = wide_format_df
                    st.session_state.data_loaded = True
                    
                    st.success("✅ Data loaded successfully!")
                    
                    # Show summary
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("👥 Students", len(students_df))
                    with col2:
                        st.metric("🏢 Companies", len(companies_df))
                    with col3:
                        st.metric("🏭 Industries", companies_df['industry'].nunique())
                    
                    # Preview
                    with st.expander("👀 Preview Your Data"):
                        st.dataframe(wide_format_df, use_container_width=True, height=300)
                    
                    # Next button
                    st.markdown("---")
                    if st.button("➡️ Continue to Review", type="primary", use_container_width=True):
                        st.session_state.step = 2
                        st.rerun()
    
    with tab2:
        st.markdown("### Create a New Template")
        st.info("💡 **Don't have a file yet?** Download our template, fill it in Excel, then upload it!")
        
        col1, col2 = st.columns(2)
        with col1:
            n_students = st.number_input("Number of Students", min_value=1, max_value=50, value=11)
        with col2:
            n_companies = st.number_input("Number of Companies", min_value=1, max_value=50, value=11)
        
        if st.button("📥 Download Template", use_container_width=True):
            template = create_template(n_students, n_companies)
            st.download_button(
                label="💾 Save Template File",
                data=template,
                file_name="placement_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

# ============================================================================
# STEP 2: REVIEW & EDIT
# ============================================================================

elif st.session_state.step == 2:
    if not st.session_state.data_loaded:
        st.warning("⚠️ Please upload data first!")
        if st.button("⬅️ Back to Step 1"):
            st.session_state.step = 1
            st.rerun()
    else:
        st.markdown("## ✏️ Step 2: Review Your Data")
        
        students_df = st.session_state.students_df
        companies_df = st.session_state.companies_df
        rankings_df = st.session_state.rankings_df
        
        # Summary cards
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("👥 Students", len(students_df))
        with col2:
            st.metric("🏢 Companies", len(companies_df))
        with col3:
            st.metric("📊 IT2 Capacity", companies_df['it2_capacity'].sum())
        with col4:
            st.metric("📊 IT3 Capacity", companies_df['it3_capacity'].sum())
        
        st.markdown("---")
        
        # Quick checks
        st.markdown("### 🔍 Quick Checks")
        is_feasible, errors, warnings = check_feasibility(students_df, companies_df)
        
        if errors:
            st.error("❌ **Problems Found:**")
            for err in errors:
                st.markdown(f"- {err}")
            st.info("💡 **Fix:** Go back and increase company capacities, or add more companies")
            
            if st.button("⬅️ Back to Fix Data"):
                st.session_state.step = 1
                st.rerun()
        else:
            st.success("✅ All checks passed! Ready to optimize!")
            
            if warnings:
                with st.expander("⚠️ Minor Warnings (Click to view)"):
                    for warn in warnings:
                        st.markdown(f"- {warn}")
        
        # Data preview
        st.markdown("---")
        st.markdown("### 📋 Your Data")
        
        tab1, tab2, tab3 = st.tabs(["🏢 Companies", "👥 Students", "📊 Rankings Summary"])
        
        with tab1:
            st.dataframe(companies_df[['company_name', 'industry', 'it2_capacity', 'it3_capacity']], 
                        use_container_width=True, height=400)
        
        with tab2:
            st.dataframe(students_df, use_container_width=True, height=400)
        
        with tab3:
            # Show top companies by average ranking
            avg_rankings = rankings_df.groupby('company_id')['ranking'].mean().sort_values(ascending=False).head(10)
            company_names = companies_df.set_index('company_id')['company_name']
            avg_rankings.index = avg_rankings.index.map(company_names)
            
            col1, col2 = st.columns(2)
            with col1:
                fig = px.bar(x=avg_rankings.values, y=avg_rankings.index, orientation='h',
                           title="🏆 Top 10 Most Wanted Companies",
                           labels={'x': 'Average Ranking', 'y': ''})
                st.plotly_chart(fig, use_container_width=True)
            
            with col2:
                # Ranking distribution
                fig = px.histogram(rankings_df[rankings_df['ranking'] > 0], x='ranking',
                                 title="📊 Distribution of Rankings",
                                 nbins=10)
                st.plotly_chart(fig, use_container_width=True)
        
        # Navigation
        st.markdown("---")
        col1, col2 = st.columns(2)
        with col1:
            if st.button("⬅️ Back to Upload", use_container_width=True):
                st.session_state.step = 1
                st.rerun()
        with col2:
            if is_feasible:
                if st.button("➡️ Optimize Placements!", type="primary", use_container_width=True):
                    st.session_state.step = 3
                    st.rerun()

# ============================================================================
# STEP 3: OPTIMIZATION & RESULTS
# ============================================================================

elif st.session_state.step == 3:
    if not st.session_state.data_loaded:
        st.warning("⚠️ Please upload data first!")
        if st.button("⬅️ Back to Start"):
            st.session_state.step = 1
            st.rerun()
    else:
        st.markdown("## 🚀 Step 3: Optimization Results")
        
        students_df = st.session_state.students_df
        companies_df = st.session_state.companies_df
        rankings_df = st.session_state.rankings_df
        
        if not st.session_state.optimization_done:
            with st.spinner("🔄 Finding optimal placements... This may take a minute..."):
                prob, x, y, rankings_dict = solve_optimization(students_df, companies_df, rankings_df)
                
                if prob.status == 1:
                    st.session_state.optimization_done = True
                    st.session_state.prob = prob
                    st.session_state.x = x
                    st.session_state.y = y
                    st.session_state.rankings_dict = rankings_dict
                else:
                    st.error("❌ Could not find a solution. Please check your data.")
                    if st.button("⬅️ Back to Review"):
                        st.session_state.step = 2
                        st.rerun()
                    st.stop()
        
        # Get results from session state
        prob = st.session_state.prob
        x = st.session_state.x
        y = st.session_state.y
        rankings_dict = st.session_state.rankings_dict
        
        st.success("✅ **Optimization Complete!**")
        
        # Metrics
        students = students_df['student_id'].tolist()
        companies = companies_df['company_id'].tolist()
        
        col1, col2, col3 = st.columns(3)
        with col1:
            obj_value = pulp.value(prob.objective)
            st.metric("🎯 Total Satisfaction Score", f"{obj_value:.0f}")
        with col2:
            max_possible = rankings_df['ranking'].max() * 2 * len(students)
            efficiency = (obj_value / max_possible * 100) if max_possible > 0 else 0
            st.metric("📈 Efficiency", f"{efficiency:.1f}%")
        with col3:
            avg_ranking = obj_value / (2 * len(students))
            st.metric("⭐ Average Ranking", f"{avg_ranking:.1f}/10")
        
        st.markdown("---")
        
        # Extract assignments
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
        
        # Results tabs
        st.markdown("### 📋 Placement Assignments")
        
        tab1, tab2, tab3 = st.tabs(["🎯 IT2 Placements", "🎯 IT3 Placements", "📊 Summary"])
        
        with tab1:
            st.dataframe(it2_df.sort_values('Student'), use_container_width=True, height=400)
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Average Ranking", f"{it2_df['Ranking'].mean():.2f}")
            with col2:
                st.metric("Highest Ranking", f"{it2_df['Ranking'].max():.0f}")
            with col3:
                st.metric("Lowest Ranking", f"{it2_df['Ranking'].min():.0f}")
        
        with tab2:
            st.dataframe(it3_df.sort_values('Student'), use_container_width=True, height=400)
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Average Ranking", f"{it3_df['Ranking'].mean():.2f}")
            with col2:
                st.metric("Highest Ranking", f"{it3_df['Ranking'].max():.0f}")
            with col3:
                st.metric("Lowest Ranking", f"{it3_df['Ranking'].min():.0f}")
        
        with tab3:
            combined = it2_df.merge(it3_df, on='Student', suffixes=('_IT2', '_IT3'))
            combined['Total_Score'] = combined['Ranking_IT2'] + combined['Ranking_IT3']
            combined = combined.sort_values('Total_Score', ascending=False)
            
            st.dataframe(combined[['Student', 'Company_IT2', 'Ranking_IT2', 'Company_IT3', 'Ranking_IT3', 'Total_Score']], 
                        use_container_width=True, height=400)
            
            # Diversity check
            st.markdown("#### ✅ Diversity Check")
            all_diverse = all(combined['Industry_IT2'] != combined['Industry_IT3'])
            if all_diverse:
                st.success("🎉 Perfect! All students have placements in different industries!")
            else:
                same_industry = combined[combined['Industry_IT2'] == combined['Industry_IT3']]
                st.warning(f"⚠️ {len(same_industry)} students have both placements in the same industry")
        
        # Charts
        st.markdown("---")
        st.markdown("### 📊 Visual Analysis")
        
        col1, col2 = st.columns(2)
        
        with col1:
            # Industry distribution
            industry_dist = pd.concat([
                it2_df['Industry'].value_counts().rename('IT2'),
                it3_df['Industry'].value_counts().rename('IT3')
            ], axis=1).fillna(0)
            
            fig = px.bar(industry_dist, barmode='group',
                       title="🏭 Placements by Industry",
                       labels={'value': 'Number of Students', 'variable': 'Iteration'})
            st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            # Ranking distribution
            all_rankings = pd.concat([
                it2_df[['Ranking']].assign(Type='IT2'),
                it3_df[['Ranking']].assign(Type='IT3')
            ])
            
            fig = px.histogram(all_rankings, x='Ranking', color='Type', barmode='overlay',
                             title="⭐ Distribution of Assigned Rankings",
                             nbins=10)
            st.plotly_chart(fig, use_container_width=True)
        
        # Export
        st.markdown("---")
        st.markdown("### 💾 Download Results")
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            it2_df.to_excel(writer, sheet_name='IT2_Placements', index=False)
            it3_df.to_excel(writer, sheet_name='IT3_Placements', index=False)
            combined.to_excel(writer, sheet_name='Combined', index=False)
            
            summary = pd.DataFrame({
                'Metric': ['Total Students', 'Total Companies', 'Satisfaction Score', 
                          'Efficiency', 'Avg IT2 Ranking', 'Avg IT3 Ranking'],
                'Value': [len(students), len(companies), f"{obj_value:.0f}",
                         f"{efficiency:.1f}%", f"{it2_df['Ranking'].mean():.2f}",
                         f"{it3_df['Ranking'].mean():.2f}"]
            })
            summary.to_excel(writer, sheet_name='Summary', index=False)
        
        st.download_button(
            label="📥 Download Complete Results (Excel)",
            data=output.getvalue(),
            file_name="placement_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )
        
        # Navigation
        st.markdown("---")
        col1, col2 = st.columns(2)
        with col1:
            if st.button("⬅️ Back to Review", use_container_width=True):
                st.session_state.step = 2
                st.rerun()
        with col2:
            if st.button("🔄 Start Over", use_container_width=True):
                # Reset everything
                st.session_state.step = 1
                st.session_state.data_loaded = False
                st.session_state.optimization_done = False
                st.rerun()

# ============================================================================
# FOOTER
# ============================================================================

st.markdown("---")
st.markdown("""
    <div style='text-align: center; color: #666; padding: 2rem;'>
        <p><b>Co-op Placement Optimizer</b> | Version 6.0 - Simplified</p>
        <p>Built with ❤️ using Streamlit & PuLP</p>
    </div>
    """, unsafe_allow_html=True)
