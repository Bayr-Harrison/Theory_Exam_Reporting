import streamlit as st
import pandas as pd
import pg8000
import os
from io import BytesIO

# Set up Streamlit
st.title("Supabase Data Exporter")
st.write("Enter the password to access the application.")

# Authentication
def authenticate(password):
    correct_password = os.environ["APP_PASSWORD"]
    return password == correct_password

# Password input
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    password = st.text_input("Password", type="password")
    if st.button("Submit"):
        if authenticate(password):
            st.session_state["authenticated"] = True
            st.success("Authenticated successfully!")
        else:
            st.error("Incorrect password. Please try again.")
else:
    # Main app functionality
    st.write("Select a date range to query the database and download results as an Excel file.")

    # Date inputs
    start_date = st.date_input("Start Date")
    end_date = st.date_input("End Date")

    # Connect to Supabase database
    def connect_to_supabase():
        return pg8000.connect(
            database=os.environ["SUPABASE_DB_NAME"],
            user=os.environ["SUPABASE_USER"],
            password=os.environ["SUPABASE_PASSWORD"],
            host=os.environ["SUPABASE_HOST"],
            port=os.environ["SUPABASE_PORT"]
        )

    # Query the database
    def query_database(start_date, end_date):
        db_query = f"""
            SELECT student_list.name,                  
                   student_list.iatc_id,
                   exam_results.nat_id,
                   student_list.class,
                   student_list.curriculum,
                   exam_results.exam,
                   exam_results.score,
                   exam_results.result,
                   exam_results.session,
                   exam_results.date,
                   exam_results.attempt_index,
                   exam_results.score_index
            FROM exam_results 
            JOIN student_list ON exam_results.nat_id = student_list.nat_id
            WHERE exam_results.date >= '{start_date}' AND exam_results.date <= '{end_date}'
            ORDER BY exam_results.date ASC, exam_results.session ASC, student_list.class, student_list.iatc_id ASC
        """
        try:
            connection = connect_to_supabase()
            cursor = connection.cursor()
            cursor.execute(db_query)
            result = cursor.fetchall()
            cursor.close()
            connection.close()
            return result
        except Exception as e:
            st.error(f"Error querying database: {e}")
            return []

    # Create and download Excel file
    def create_excel(data):
        col_names = [
            'Name', 'IATC ID', 'National ID', 'Class', 'Faculty', 
            'Exam', 'Score', 'Result', 'Session', 'Date', 
            'Attempt Index', 'Score Index'
        ]
        df = pd.DataFrame(data, columns=col_names)
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Results')
        output.seek(0)
        return output

    # Button to trigger query and download
    if st.button("Query and Download"):
        if start_date and end_date:
            st.write("Querying database...")
            data = query_database(start_date, end_date)
            if data:
                st.write(f"Query returned {len(data)} rows.")
                excel_file = create_excel(data)
                st.download_button(
                    label="Download Excel File",
                    data=excel_file,
                    file_name="query_results.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("No data found for the specified date range.")
        else:
            st.error("Please select both start and end dates.")
