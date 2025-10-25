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
