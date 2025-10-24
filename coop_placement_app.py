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
                        ["Data Setup", "Exploratory Analysis", "Clustering", "Optimization"],
                        index=0)

# Helper functions
def generate_synthetic_data(n_students=15):
    """Generate synthetic dataset with 15 students"""
    np.random.seed(42)
    
    # Students
    student_names = [f"Student_{i+1:02d}" for i in range(n_students)]
    students_df = pd.DataFrame({
        'student_id': range(1, n_students+1),
        'student_name': student_names,
        'gpa': np.random.uniform(3.0, 4.0, n_students).round(2)
    })
    
    # Companies - exactly 15 to match students
    company_info = [
        ("QBE", "General Insurance", 1, 1),
        ("IAG", "General Insurance", 1, 1),
        ("Suncorp", "General Insurance", 1, 1),
        ("Allianz", "General Insurance", 1, 1),
        ("Deloitte", "Consultancy", 1, 1),
        ("PwC", "Consultancy", 1, 1),
        ("KPMG", "Consultancy", 1, 1),
        ("EY", "Consultancy", 1, 1),
        ("AMP", "Life Insurance", 1, 1),
        ("MLC", "Life Insurance", 1, 1),
        ("TAL", "Life Insurance", 1, 1),
        ("Zurich", "Life Insurance", 1, 1),
        ("NDIS Provider A", "Care/Disability", 1, 1),
        ("NDIS Provider B", "Care/Disability", 1, 1),
        ("NDIS Provider C", "Care/Disability", 1, 1),
    ]
    
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

def create_excel_template():
    """Create Excel template for data input"""
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
    
    for i in range(2, 17):
        ws_students[f'A{i}'] = i-1
        ws_students[f'B{i}'] = f'Student_{i-1:02d}'
        ws_students[f'C{i}'] = ''
    
    # Companies sheet
    ws_companies = wb.create_sheet("Companies")
    headers = ['company_id', 'company_name', 'industry', 'it2_capacity', 'it3_capacity']
    for col, header in enumerate(headers, 1):
        cell = ws_companies.cell(1, col, header)
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill(start_color='366092', fill_type='solid')
    
    industries = ['General Insurance', 'Consultancy', 'Life Insurance', 'Care/Disability']
    for i in range(2, 17):
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
    ws_rankings['E2'] = '1. Fill in student GPAs in Students sheet'
    ws_rankings['E3'] = '2. Update company names and industries in Companies sheet'
    ws_rankings['E4'] = '3. Enter rankings (1-10, higher is better) for each student-company pair'
    ws_rankings['E5'] = '4. Each student should rank all companies'
    
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
    
    if len(students_df) != len(companies_df):
        errors.append(f"Number of students ({len(students_df)}) must equal number of companies ({len(companies_df)})")
    
    expected_rankings = len(students_df) * len(companies_df)
    if len(rankings_df) != expected_rankings:
        errors.append(f"Expected {expected_rankings} rankings, got {len(rankings_df)}")
    
    if rankings_df['ranking'].min() < 1 or rankings_df['ranking'].max() > 10:
        errors.append("Rankings must be between 1 and 10")
    
    return errors

# PAGE 1: DATA SETUP
if page == "Data Setup":
    st.title("📊 Data Setup")
    
    tab1, tab2, tab3 = st.tabs(["Synthetic Data", "Manual Input", "Excel Upload"])
    
    # Tab 1: Synthetic Data
    with tab1:
        st.header("Generate Synthetic Dataset")
        st.write("Click the button below to generate a synthetic dataset with 15 students and 15 companies.")
        
        if st.button("Generate Synthetic Data", type="primary"):
            students_df, companies_df, rankings_df = generate_synthetic_data()
            st.session_state.students_df = students_df
            st.session_state.companies_df = companies_df
            st.session_state.rankings_df = rankings_df
            st.session_state.data_loaded = True
            st.success("✅ Synthetic data generated successfully!")
        
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
    
    # Tab 2: Manual Input
    with tab2:
        st.header("Manual Data Input")
        st.warning("⚠️ Manual input is for demonstration. For 15 students × 15 companies = 225 rankings, use Excel upload instead.")
        
        st.write("**Note:** This simplified form allows you to set a few rankings. Use Excel for complete data entry.")
        
        if st.session_state.data_loaded:
            st.write("### Edit Rankings")
            student_id = st.selectbox("Select Student", st.session_state.students_df['student_id'].tolist())
            company_id = st.selectbox("Select Company", st.session_state.companies_df['company_id'].tolist())
            ranking = st.slider("Ranking (1-10, higher = more preferred)", 1, 10, 5)
            
            if st.button("Update Ranking"):
                mask = (st.session_state.rankings_df['student_id'] == student_id) & \
                       (st.session_state.rankings_df['company_id'] == company_id)
                st.session_state.rankings_df.loc[mask, 'ranking'] = ranking
                st.success(f"✅ Updated ranking for Student {student_id} - Company {company_id}")
        else:
            st.info("Generate synthetic data first or upload Excel file.")
    
    # Tab 3: Excel Upload
    with tab3:
        st.header("Excel Data Upload")
        
        # Download template
        st.subheader("Step 1: Download Template")
        template_buffer = create_excel_template()
        st.download_button(
            label="📥 Download Excel Template",
            data=template_buffer,
            file_name="coop_placement_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # Upload file
        st.subheader("Step 2: Upload Completed File")
        uploaded_file = st.file_uploader("Upload Excel file", type=['xlsx'])
        
        if uploaded_file:
            students_df, companies_df, rankings_df, error = load_excel_data(uploaded_file)
            
            if error:
                st.error(f"❌ Error loading file: {error}")
            else:
                errors = validate_data(students_df, companies_df, rankings_df)
                
                if errors:
                    st.error("❌ Data validation failed:")
                    for err in errors:
                        st.write(f"- {err}")
                else:
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

# PAGE 2: EXPLORATORY ANALYSIS
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
        
        fig = px.bar(avg_rankings, x='company_name', y='ranking', color='industry',
                    title="Average Ranking by Company",
                    labels={'ranking': 'Average Ranking', 'company_name': 'Company'})
        fig.update_xaxes(tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)
        
        # Heatmap of rankings
        st.header("Student-Company Ranking Heatmap")
        pivot_rankings = rankings_df.pivot(index='student_id', columns='company_id', values='ranking')
        
        fig = px.imshow(pivot_rankings, 
                       labels=dict(x="Company ID", y="Student ID", color="Ranking"),
                       title="Student Preferences Heatmap",
                       aspect="auto",
                       color_continuous_scale='RdYlGn')
        st.plotly_chart(fig, use_container_width=True)

# PAGE 3: CLUSTERING
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
        
        # Create distance matrix
        pivot_rankings = rankings_df.pivot(index='student_id', columns='company_id', values='ranking')
        
        # Compute linkage
        linkage_methods = ['ward', 'complete', 'average', 'single']
        method = st.selectbox("Select Linkage Method", linkage_methods, index=0)
        
        Z = linkage(pivot_rankings, method=method)
        
        # Dendrogram
        st.header("Dendrogram")
        fig = go.Figure()
        
        # Calculate dendrogram
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
        
        # Select number of clusters
        st.header("Cluster Assignment")
        n_clusters = st.slider("Select Number of Clusters", 2, 10, 3)
        
        # Get cluster assignments
        clusters = fcluster(Z, n_clusters, criterion='maxclust')
        
        # Add clusters to students dataframe
        students_clustered = students_df.copy()
        students_clustered['cluster'] = clusters
        
        # Display cluster information
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
        
        # Cluster characteristics
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

# PAGE 4: OPTIMIZATION
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
        
        if st.button("🚀 Solve Optimization Problem", type="primary"):
            with st.spinner("Solving optimization problem..."):
                # Create the optimization problem
                prob = pulp.LpProblem("CoopPlacement", pulp.LpMaximize)
                
                # Decision variables
                students = students_df['student_id'].tolist()
                companies = companies_df['company_id'].tolist()
                
                x = {}  # IT2 assignments
                y = {}  # IT3 assignments
                
                for i in students:
                    for j in companies:
                        x[i, j] = pulp.LpVariable(f"x_{i}_{j}", cat='Binary')
                        y[i, j] = pulp.LpVariable(f"y_{i}_{j}", cat='Binary')
                
                # Objective function
                rankings_dict = rankings_df.set_index(['student_id', 'company_id'])['ranking'].to_dict()
                prob += pulp.lpSum([rankings_dict[i, j] * (x[i, j] + y[i, j]) 
                                   for i in students for j in companies])
                
                # Constraints
                # 1. Each student does one IT2
                for i in students:
                    prob += pulp.lpSum([x[i, j] for j in companies]) == 1
                
                # 2. Each student does one IT3
                for i in students:
                    prob += pulp.lpSum([y[i, j] for j in companies]) == 1
                
                # 3. IT2 capacity constraints
                for j in companies:
                    capacity = companies_df[companies_df['company_id'] == j]['it2_capacity'].values[0]
                    prob += pulp.lpSum([x[i, j] for i in students]) <= capacity
                
                # 4. IT3 capacity constraints
                for j in companies:
                    capacity = companies_df[companies_df['company_id'] == j]['it3_capacity'].values[0]
                    prob += pulp.lpSum([y[i, j] for i in students]) <= capacity
                
                # 5. IT2 and IT3 in different companies
                for i in students:
                    for j in companies:
                        prob += x[i, j] + y[i, j] <= 1
                
                # 6. Industry diversity constraints
                industry_companies = companies_df.groupby('industry')['company_id'].apply(list).to_dict()
                
                for i in students:
                    for industry, company_list in industry_companies.items():
                        prob += pulp.lpSum([x[i, j] + y[i, j] for j in company_list]) <= 1
                
                # Solve
                solver = pulp.PULP_CBC_CMD(msg=0)
                prob.solve(solver)
                
                # Display results
                st.success(f"✅ Optimization Status: {pulp.LpStatus[prob.status]}")
                
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("Objective Value", f"{pulp.value(prob.objective):.2f}")
                with col2:
                    st.metric("Max Possible", f"{rankings_df['ranking'].max() * 2 * len(students)}")
                
                # Extract solution
                if prob.status == 1:  # Optimal
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
                    
                    # Display assignments
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
                    
                    # Visualization
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
                    
                    # Download results
                    st.header("Export Results")
                    
                    # Create Excel file
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

# Footer
st.sidebar.markdown("---")
st.sidebar.info("""
**Co-op Placement Optimizer**  
Version 1.0  
Built with Streamlit & PuLP
""")
