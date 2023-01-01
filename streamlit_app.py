import base64

import streamlit as st
import pandas as pd
#import pyodbc
import numpy as np
import time
import plotly.express as px  # pip install plotly-express
import  openpyxl

try:
    st.set_page_config(layout='wide', initial_sidebar_state='expanded')

    with open('style.css') as f:
        st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)


    # ---- READ ----

    #def sql_connection():
        #database = pyodbc.connect('Driver={SQL Server};'
                                  #'Server=HAFIDA\SQLSERVER19;'
                                  #'Database=FDB;'
                                  #'Trusted_Connection=yes;')
        #return database


    #cursor = sql_connection().cursor()


    #def get_data_from_db():
        # Executing a SQL Query
        #query = "SELECT * FROM Employees"
        #df = pd.read_sql(query, sql_connection())
        #return df


    #df = get_data_from_db()
    def get_data_from_excel():
        df = pd.read_excel(
        io="Employees.xlsx",
        engine="openpyxl",
        sheet_name="Sheet1"
    )
        return df
    df = get_data_from_excel()

    df2 = pd.DataFrame(df['HireDate'].dt.year.value_counts().reset_index().values, columns=["Year", "Hired"])
    df2 = df2.sort_values(by=['Year'])

    # dashboard title
    st.title("Real-Time / Live Data Dashboard")

    # ---- SIDEBAR ----

    # avg_sal =
    st.sidebar.image("learning-and-development-strategies.png", use_column_width=True)

    st.sidebar.header('Filters')

    st.sidebar.subheader('Select one filter at a time')

    department = st.sidebar.multiselect("Select the Department", pd.unique(df["Department"]))

    business = st.sidebar.multiselect("Select the Business Unit", pd.unique(df["BusinessUnit"]))

    country = st.sidebar.multiselect("Select the Country:", pd.unique(df["Country"]))

    st.sidebar.markdown('''
  ---
  Created by Hidani Hafida
  ''')

    # dataframe filter
    df_selected = df.query("(`Department`==@department) or (`BusinessUnit`==@business) or (`Country`==@country)")

    # -------- MAIN ---------

    # creating a single-element container
    placeholder = st.empty()

    # Row B

    # fill in those three columns with respective metrics or KPIs
    # creating KPIs

    # near real-time / live feed simulation
    for seconds in range(200):
        # creating KPIs
        avg_age = np.mean(df_selected["Age"])
        count_females = int(
            df_selected[(df_selected["Gender"] == "F")]["Gender"].count()

        )
        count_males = int(
            df_selected[(df_selected["Gender"] == "M")]["Gender"].count()

        )

        with placeholder.container():
            # Row A #  Some Basic information about data
            st.markdown('### General Details')
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Employees", len(df[df.Exited == False]))
            col2.metric("Department", str(df.Department.unique().size))
            col3.metric("Business Unit", str(df.BusinessUnit.unique().size))
            col4.metric("Exited", ~ len(df[df.Exited == True]))

            # Row B# # 2 Checkbox to show dataset
            if st.checkbox("Show Data"):
                st.dataframe(df)

                def filedownload(df):
                    csv = df.to_csv(index=False)
                    b64 = base64.b64encode(csv.encode()).decode()  # strings <-> bytes conversions
                    href = f'<a href="data:file/csv;base64,{b64}" download="Employees.csv">Download CSV File</a>'
                    return href


                st.markdown(filedownload(df), unsafe_allow_html=True)


            # Row C# # 3 Total data data distribution according to Platform or Genre accornding to the selection
            op = st.selectbox("Select one of the following", ["JobTitle","Ethnicity"])
            st.subheader(op + " wise data distribution")
            st.bar_chart(df[op].value_counts(), height=400, use_container_width=True,
                                     )

            c1,c2=st.columns(2)

            with c1:
                st.markdown("#### Gender Overview")
                fig2 = px.pie(df, names='Gender',
                              hole=0.7,height=400,width=400,
                              color_discrete_sequence=['#A3E0F2', '#3C6BAD', '#FFFFFF'])
                fig2.update_layout(plot_bgcolor='#FFFFFF', paper_bgcolor='#FFFFFF')
                st.write(fig2)

            with c2:
                st.markdown("#### Top job titles")
                job_title = df.groupby('JobTitle').size().reset_index().sort_values(by=0, ascending=False)
                job_title.columns = ['JobTitle', 'count']

                fig5 = px.bar(job_title, x='JobTitle', y='count', color='JobTitle', text='count',
                              color_discrete_sequence=['#CFE4FF', '#F47269', '#3C6BAD', '#ECECEE', '#A3E0F2', '#FEBBB7',
                                                       '#4475B7'])

                fig5.update_layout(plot_bgcolor='#FFFFFF', paper_bgcolor='#FFFFFF')
                st.write(fig5)





            # Row D
            left, middle, right = st.columns((2, 5, 2))
            with middle:
                st.markdown("#### Hires by Year")
                fig4 = px.line(
                    df2,
                    x="Year",
                    y="Hired",
                    template="plotly_white",
                    color_discrete_sequence=['#3C6BAD']
                )
                fig4.update_layout(plot_bgcolor='#FFFFFF', paper_bgcolor='#FFFFFF')
                st.write(fig4)



            # create two columns for charts

            c3,c4=st.columns(2)

            with c3:
                st.markdown("#### Employee by Department")
                dept = df.groupby('Department').size().reset_index().sort_values(by=0, ascending=False)
                dept.columns = ['Department', 'count']
                fig = px.bar(dept, x='Department', y='count', color='Department',
                             color_discrete_sequence=['#CFE4FF', '#F47269', '#3C6BAD', '#ECECEE', '#A3E0F2', '#FEBBB7',
                                                      '#4475B7']
                             )
                fig.update_layout(plot_bgcolor='#FFFFFF', paper_bgcolor='#FFFFFF')
                st.write(fig)

            with c4:
                st.markdown("#### Employee by City")
                city = df.groupby('City').size().reset_index().sort_values(by=0, ascending=False)
                city.columns = ['City', 'count']
                fig3 = px.bar(city, x='City', y='count', text='City',
                              color_discrete_sequence=['#F47269', '#3C6BAD', '#ECECEE', '#A3E0F2', '#FEBBB7',
                                                       '#4475B7'])
                fig3.update_traces(texttemplate='%{text:.2s}', textposition='outside')
                fig3.update_layout(uniformtext_minsize=8, uniformtext_mode='hide')
                fig3.update_layout(plot_bgcolor='#FFFFFF', paper_bgcolor='#FFFFFF')
                st.write(fig3)









            # Row E

            st.markdown('### Filtered Details')
            # create three columns
            kpi1, kpi2, kpi3 = st.columns(3)

            # fill in those three columns with respective metrics or KPIs
            kpi1.metric(
                label="Age",
                value=round(avg_age)
            )

            kpi2.metric(
                label="Females ",
                value=int(count_females)
            )

            kpi3.metric(
                label="Males",
                value=int(count_males)
            )

            # Row F
            st.dataframe(df_selected)

            # Row

            c5,c6 = st.columns(2)

            with c5:
                st.markdown('### Gender Distribution')
                by_gender = df_selected.groupby('Gender').size().reset_index().sort_values(by=0, ascending=False)
                by_gender.columns = ['Gender', 'count']

                figg = px.pie(by_gender, values='count', names='Gender',
                              color_discrete_sequence=['#F47269', '#3C6BAD', '#ECECEE', '#A3E0F2', '#FEBBB7',
                                                       '#4475B7'])

                figg.update_layout(plot_bgcolor='#FFFFFF', paper_bgcolor='#FFFFFF')
                st.write(figg)

            with c6:
                st.markdown('### Salary Average by Gender')
                avg_salary = df_selected.groupby(['Gender'],as_index=False).Salary.mean()
                avg_salary.columns = ['Gender', 'avgS']
                figg = px.bar(avg_salary, x='Gender', y='avgS',
                              color_discrete_sequence=['#A3E0F2', '#FEBBB7',
                                                       '#4475B7'])
                figg.update_layout(uniformtext_minsize=8, uniformtext_mode='hide')
                figg.update_layout(plot_bgcolor='#FFFFFF', paper_bgcolor='#FFFFFF')
                st.write(figg)





            time.sleep(1)



except:
  # Prevent the error from propagating into your Streamlit app.
  pass

        #








