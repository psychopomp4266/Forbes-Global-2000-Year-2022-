import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
import numpy as np
import openpyxl

# Load the dataset
data = pd.read_excel('Forbes Global 2000 (Year 2022).xlsx')

# Clean and preprocess the data if needed

# Perform exploratory data analysis
def perform_eda():
    st.header('Exploratory Data Analysis')
    
    # Display basic information about the dataset
    st.subheader(':blue[Dataset Information] ')
    st.write(f"Shape: {data.shape}")
    st.write(f"Columns: {', '.join(data.columns)}")

# Display the first few rows of the dataset
    st.subheader(':blue[Head]')
    st.write(data.head())

    # Display summary statistics
    st.subheader(' :green[Summary Statistics]')
    st.write(data.describe())
    
    # Visualize insights
    st.subheader('Data Visualizations')
    
    # Select a visualization option
    visualization_option = st.selectbox('Select a visualization option', ['Sales Distribution', 'Profits Distribution', 'Assets Distribution', 'Market Value Distribution', 'Top Companies by Industry', 'Top Companies by Country'])

    # Set the option to avoid deprecation warning for matplotlib.pyplot
    st.set_option('deprecation.showPyplotGlobalUse', False)
    st.caption('Scroll up/down on the plot to zoom in/out of line or bar plots')
    # Plot the selected visualization
    if visualization_option == 'Sales Distribution':
        st.write('Sales Distribution')
        plot_type = st.selectbox('Select a plot type', ['Line Plot', 'Bar Plot', 'Histogram'])
        if plot_type == 'Line Plot':
            st.line_chart(data['Sales'])
        elif plot_type == 'Bar Plot':
            st.bar_chart(data['Sales'])
        elif plot_type == 'Histogram':
            sns.histplot(data['Sales'])
            st.pyplot()
    elif visualization_option == 'Profits Distribution':
        st.write('Profits Distribution')
        plot_type = st.selectbox('Select a plot type', ['Line Plot', 'Bar Plot', 'Histogram'])
        if plot_type == 'Line Plot':
            st.line_chart(data['Profits'])
        elif plot_type == 'Bar Plot':
            st.bar_chart(data['Profits'])
        elif plot_type == 'Histogram':
            sns.histplot(data['Profits'])
            st.pyplot()
    elif visualization_option == 'Assets Distribution':
        st.write('Assets Distribution')
        plot_type = st.selectbox('Select a plot type', ['Line Plot', 'Bar Plot', 'Histogram'])
        if plot_type == 'Line Plot':
            st.line_chart(data['Assets'])
        elif plot_type == 'Bar Plot':
            st.bar_chart(data['Assets'])
        elif plot_type == 'Histogram':
            sns.histplot(data['Assets'])
            st.pyplot()
    elif visualization_option == 'Market Value Distribution':
        st.write('Market Value Distribution')
        plot_type = st.selectbox('Select a plot type', ['Line Plot', 'Bar Plot', 'Histogram'])
        if plot_type == 'Line Plot':
            st.line_chart(data['Market_Value'])
        elif plot_type == 'Bar Plot':
            st.bar_chart(data['Market_Value'])
        elif plot_type == 'Histogram':
            sns.histplot(data['Market_Value'])
            st.pyplot()
    elif visualization_option == 'Top Companies by Industry':
        st.write('Top Companies by Industry')
        industry_counts = data['Industry'].value_counts().head(10)
        plot_type = st.selectbox('Select a plot type', ['Bar Plot', 'Pie Chart'])
        if plot_type == 'Bar Plot':
            plt.figure(figsize=(10, 6))
            sns.barplot(x=industry_counts.values, y=industry_counts.index)
            plt.xlabel('Number of Companies')
            plt.ylabel('Industry')
            plt.title('Top 10 Industries with Most Companies')
            st.pyplot()
        elif plot_type == 'Pie Chart':
            plt.figure(figsize=(10, 6))
            plt.pie(industry_counts.values, labels=industry_counts.index, autopct='%1.1f%%')
            plt.title('Top 10 Industries with Most Companies')
            st.pyplot()
    elif visualization_option == 'Top Companies by Country':
        st.write('Top Companies by Country')
        country_counts = data['Country'].value_counts().head(10)
        plot_type = st.selectbox('Select a plot type', ['Bar Plot', 'Pie Chart'])
        if plot_type == 'Bar Plot':
            plt.figure(figsize=(10, 6))
            sns.barplot(x=country_counts.values, y=country_counts.index)
            plt.xlabel('Number of Companies')
            plt.ylabel('Country')
            plt.title('Top 10 Countries with Most Companies')
            st.pyplot()
        elif plot_type == 'Pie Chart':
            plt.figure(figsize=(10, 6))
            plt.pie(country_counts.values, labels=country_counts.index, autopct='%1.1f%%')
            plt.title('Top 10 Countries with Most Companies')
            st.pyplot()
# Build the Streamlit web app
def main():
    st.title('Forbes Global 2000 Insights')
    st.write('Displaying insights of the companies listed in the Forbes Global 2000.')

    # Sidebar options
    option = st.sidebar.selectbox('Select an option', ['Exploratory Data Analysis', 'Insights'])

    if option == 'Exploratory Data Analysis':
        perform_eda()
    elif option == 'Insights':
        st.header('Insights')
        st.write('Choose from the following visualizations to gain further insights:')

        # Additional insights and visualizations
        insights_options = ['Correlation Analysis', 'Sector-wise Analysis', 'Country-wise Analysis', 'Top Performers',
                            'Geographic Analysis',  'Comparative Analysis',
                            'Market Capitalization Analysis', 'Company Rank Analysis']
        
        chosen_insights = st.multiselect('Select Insights', insights_options)

        if 'Correlation Analysis' in chosen_insights:
            # Perform correlation analysis and visualize using a heatmap
            st.subheader('Correlation Analysis')
            correlation = data[['Sales', 'Profits', 'Assets', 'Market_Value']].corr()
            plt.figure(figsize=(10, 8))
            sns.heatmap(correlation, annot=True, cmap='coolwarm', linewidths=0.5)
            st.pyplot()

        if 'Sector-wise Analysis' in chosen_insights:
            # Perform sector-wise analysis and visualize using a bar chart or pie chart
            st.subheader('Sector-wise Analysis')
            sector_counts = data['Industry'].value_counts()
            plot_type = st.selectbox('Select a plot type', ['Bar Plot', 'Pie Chart'])
            if plot_type == 'Bar Plot':
                plt.figure(figsize=(10, 7))
                sns.barplot(x=sector_counts.values, y=sector_counts.index)
                plt.xlabel('Number of Companies')
                plt.ylabel('Industry')
                plt.title('Companies by Industry')
                st.pyplot()
            elif plot_type == 'Pie Chart':
                plt.figure(figsize=(15, 10))
                plt.pie(sector_counts.values, labels=sector_counts.index, autopct='%1.1f%%')
                plt.title('Companies by Industry')
                st.pyplot()

        if 'Country-wise Analysis' in chosen_insights:
            # Perform country-wise analysis and visualize using a bar chart or pie chart
            st.subheader('Country-wise Analysis')
            country_counts = data['Country'].value_counts()
            plot_type = st.selectbox('Select a plot type', ['Bar Plot', 'Pie Chart'])
            if plot_type == 'Bar Plot':
                plt.figure(figsize=(10, 12))
                sns.barplot(x=country_counts.values, y=country_counts.index)
                plt.xlabel('Number of Companies')
                plt.ylabel('Country')
                plt.title('Companies by Country')
                st.pyplot()
            elif plot_type == 'Pie Chart':
                plt.figure(figsize=(20, 9))
                plt.pie(country_counts.values, labels=country_counts.index, autopct='%1.1f%%')
                plt.title('Companies by Country')
                st.pyplot()

        if 'Top Performers' in chosen_insights:
            # Display top performers based on different metrics
            st.subheader('Top Performers')
            st.write('Choose from the following performance grading metrics:')

            # Additional insights and visualizations
            performer_options = ['Sales', 'Profits', 'Assets', 'Market Value']
        
            chosen_options = st.multiselect('Select Top 10 of:', performer_options)

            top_sales = data.sort_values('Sales', ascending=False).head(10)
            top_profits = data.sort_values('Profits', ascending=False).head(10)
            top_assets = data.sort_values('Assets', ascending=False).head(10)
            top_market_value = data.sort_values('Market_Value', ascending=False).head(10)
            
            if 'Sales' in chosen_options:
                st.write(':blue[Top 10 Companies by Sales]')
                st.table(top_sales[['Rank_nr', 'Company', 'Sales']])
            
            if 'Profits' in chosen_options:
                st.write(':blue[Top 10 Companies by Profits]')
                st.table(top_profits[['Rank_nr', 'Company', 'Profits']])
            
            if 'Assets' in chosen_options:
                st.write(':blue[Top 10 Companies by Assets]')
                st.table(top_assets[['Rank_nr', 'Company', 'Assets']])
            
            
            if 'Market Value' in chosen_options:
                st.write(':blue[Top 10 Companies by Market Value]')
                st.table(top_market_value[['Rank_nr', 'Company', 'Market_Value']])

        
        if 'Geographic Analysis' in chosen_insights:
        # Perform geographic analysis and visualize using a scatter plot or bubble chart
            st.subheader('Geographic Analysis')
            st.write('Choose a metric to represent on the map:')
            metric = st.radio('Metric', ('Sales', 'Profits', 'Assets'))
            
            # Perform geographic analysis based on the chosen metric and ignore negative values
            if metric == 'Sales':
                metric_data = np.maximum(data['Sales'], 0)
            elif metric == 'Profits':
                metric_data = np.maximum(data['Profits'], 0)
            elif metric == 'Assets':
                metric_data = np.maximum(data['Assets'], 0)
            
            fig = px.scatter_geo(data, locations='Country', locationmode='country names', color=metric_data,
                                    size=metric_data, hover_data=['Company', 'Country', metric, 'Market_Value'])
            st.plotly_chart(fig)


                
        if 'Comparative Analysis' in chosen_insights:
            # Perform comparative analysis and visualize using a bar plot or box plot
            st.subheader('Comparative Analysis')
            st.write('Select metrics to compare:')
            
            try:
                selected_metrics = st.multiselect('Metrics', ['Sales', 'Profits', 'Assets', 'Market_Value'])
                comparative_data = data[selected_metrics]

                plot_type = st.selectbox('Select a plot type', ['Bar Plot', 'Box Plot'])
                
                if plot_type == 'Bar Plot':
                    plt.figure(figsize=(10, 6))
                    comparative_data.plot(kind='bar')
                    plt.xlabel('Company')
                    plt.ylabel('Metric Value')
                    plt.title('Comparative Analysis')
                    plt.legend(loc='upper left')
                    st.pyplot()
                
                elif plot_type == 'Box Plot':
                    try:
                        plt.figure(figsize=(10, 6))
                        comparative_data.boxplot()
                        plt.xlabel('Metric')
                        plt.ylabel('Metric Value')
                        plt.title('Comparative Analysis')
                        st.pyplot()
                    except ValueError:
                        st.error("Box plot requires numerical columns, nothing to plot.")
                        st.error("Please choose a metric.")        
                
                
            except TypeError:
                st.error("Please choose at least one of the metrics to make the plot visible.")
            
            
        if 'Market Capitalization Analysis' in chosen_insights:
            # Perform market capitalization analysis and visualize using a scatter plot or bar plot
            st.subheader('Market Capitalization Analysis')
            market_capitalization = data['Market_Value']

            plot_type = st.selectbox('Select a plot type', ['All Companies', 'Top 10 Companies'])
            if plot_type == 'All Companies':
                plt.figure(figsize=(10, 6))
                sns.barplot(x=data['Rank_nr'], y=market_capitalization)
                plt.xlabel('Company Rank')
                plt.ylabel('Market Capitalization')
                plt.title('Market Capitalization Analysis')
                st.pyplot()
            elif plot_type == 'Top 10 Companies':
                plt.figure(figsize=(10, 6))
                sns.barplot(x=data['Rank_nr'], y=market_capitalization.nlargest(10))
                plt.xlabel('Company Rank')
                plt.ylabel('Market Capitalization')
                plt.title('Market Capitalization Analysis')
                st.pyplot()


        if 'Company Rank Analysis' in chosen_insights:
            # Perform company rank analysis and visualize using a bar plot or violin plot
            st.subheader('Company Rank Analysis')
            st.write('Select metrics to analyze by company rank:')
            selected_metrics = st.multiselect('Metrics', ['Sales', 'Profits', 'Assets', 'Market_Value'])
            rank_data = data[['Rank_nr'] + selected_metrics]

            plot_type = st.selectbox('Select a plot type', ['Bar Plot', 'Violin Plot'])
            
            try:
                if plot_type == 'Bar Plot':
                    rank_data.set_index('Rank_nr', inplace=True)
                    plt.figure(figsize=(10, 6))
                    rank_data.plot(kind='bar', stacked=False)
                    plt.xlabel('Company Rank')
                    plt.ylabel('Metric Value')
                    plt.title('Company Rank Analysis')
                    plt.legend(loc='upper left')
                    st.pyplot()
                elif plot_type == 'Violin Plot':
                    plt.figure(figsize=(10, 6))
                    melted_data = pd.melt(rank_data, id_vars=['Rank_nr'], value_vars=selected_metrics,
                                        var_name='Metric', value_name='Metric Value')
                    sns.violinplot(x='Rank_nr', y='Metric Value', data=melted_data)
                    plt.xlabel('Company Rank')
                    plt.ylabel('Metric Value')
                    plt.title('Company Rank Analysis')
                    st.pyplot()

            except TypeError:
                st.error("Please choose at least one of the metrics to make the plot visible.")
            

if __name__ == '__main__':
    main()
