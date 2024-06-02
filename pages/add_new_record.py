import streamlit as st
import pandas as pd
import os
from App import load_dataframe

home_page      = st.sidebar.page_link('App.py',                  label="HOME",           icon="🏡")
record_page    = st.sidebar.page_link('pages/record.py',         label="RECORD",         icon="📓")    
add_new_record = st.sidebar.page_link('pages/add_new_record.py', label="ADD NEW RECORD", icon="✒️") 

file_path = os.path.join(os.getcwd(), 'InvoiceLogTemplate.xlsx')  # Full file path
worksheet_name = "Clients"
df = load_dataframe(file_path, worksheet_name)


tab1, tab2 = st.tabs(['Add New Client', 'Add Project'])
with tab1:
    new_client  = st.text_input("Client",  key="new_client")
    new_project = st.text_input("Project", key="project_name")

with tab2:
    clients_drop_down  = df['Client'].unique()
    selected_client   = st.selectbox("Select Client", clients_drop_down)
    new_project_existing_client = st.text_input("Project", key="project_name_for_existing_client")
    
col1, col2, col3 = st.columns([1,1,2])
with col1:
    save_records   = st.button("Update Record", key="update_record")
with col2:    
    display_record = st.button("Display Record")

try:
    if save_records:  
        if not ((df['Client'] == new_client) & (df['Project'] == new_project)).any():
            if new_client or new_project:
                add_new_record = {
                    'Client' : new_client,
                    'Project': new_project,
                }
                # Read existing data from the Excel file
                xl = pd.ExcelFile(file_path)
                # Load all sheets into a dictionary of DataFrames
                dfs = {sheet_name: xl.parse(sheet_name) for sheet_name in xl.sheet_names}
                # Update the specific sheet with the new record
                if worksheet_name in dfs:
                    df = dfs[worksheet_name]
                    df = pd.concat([df, pd.DataFrame([add_new_record])], ignore_index=True)
                    dfs[worksheet_name] = df
                else:
                    # Handle case where worksheet_name doesn't exist
                    dfs[worksheet_name] = pd.DataFrame([add_new_record])
                # Write all sheets back to the Excel file
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    for sheet_name, df in dfs.items():
                        df.to_excel(writer, index=False, sheet_name=sheet_name)

            st.success("Record saved successfully.")   


        elif not ((df['Client'] == selected_client) & (df['Project'] == new_project_existing_client)).any():
            if selected_client and new_project_existing_client:
                add_new_record = {
                    'Client' : selected_client,
                    'Project': new_project_existing_client,
                }
                # Read existing data from the Excel file
                xl = pd.ExcelFile(file_path)
                # Load all sheets into a dictionary of DataFrames
                dfs = {sheet_name: xl.parse(sheet_name) for sheet_name in xl.sheet_names}
                # Update the specific sheet with the new record
                if worksheet_name in dfs:
                    df = dfs[worksheet_name]
                    df = pd.concat([df, pd.DataFrame([add_new_record])], ignore_index=True)
                    dfs[worksheet_name] = df
                else:
                    # Handle case where worksheet_name doesn't exist
                    dfs[worksheet_name] = pd.DataFrame([add_new_record])
                # Write all sheets back to the Excel file
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    for sheet_name, df in dfs.items():
                        df.to_excel(writer, index=False, sheet_name=sheet_name)

            st.success("Record saved successfully.")   

        else:
            st.warning("Record Already Exist")
except:
    pass

if display_record:
    st.dataframe(df)
