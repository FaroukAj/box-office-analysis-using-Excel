# box-office-analysis-using-Excel

# 🎬 Movie Data Analysis Project (Excel)

This project explores a dataset containing information on thousands of movies, originally sourced from an NBC Universal data competition. The goal was to clean the data and use **Microsoft Excel** to answer business-focused questions related to movie revenue, profit, ratings, and production company performance.

## 📊 Tools Used

- Microsoft Excel

## 📁 Dataset Overview

The dataset includes variables such as:
- Movie title
- Release date
- Genre
- IMDB rating
- Metacritic rating
- Budget
- Box Office Gross
- Production company
- IMDB votes

## 🧹 Step 1: Data Cleaning (Excel)

I cleaned the dataset using Power Query by removing rows with:
- Missing or invalid values in key columns like `imdb_rating`, `metacritic`, `imdb_votes`, `release_date`, `genre`, `Budget`, and `Box Office Gross`
- Values of `0` or text in the `Budget` and `Box Office Gross` columns were considered invalid

After cleaning, the worksheet was renamed to **Clean Data** and saved. The cleaned dataset included approximately **2,282 rows**.

## 📊 Step 2: Excel Data Insights

### Top 10 Companies by Total Box Office Gross
- I used a PivotTable to sum `Box Office Gross` by `Production Company`
- Sorted the values to display the top 10 highest-grossing studios
- Worksheet renamed to **Question 3.1**

### Top 10 Companies by Number of Movies
- Created a PivotTable to count the number of movies per production company
- Sorted the results to show the top 10 studios by volume
- Worksheet renamed to **Question 3.2**

### Top 10 Companies by Average Box Office Per Movie
- Built a PivotTable showing total revenue and count of movies per studio
- Calculated the average revenue per movie and sorted for the top 10
- Worksheet renamed to **Question 3.3**

### Treemap Chart
- I visualized the `Box Office Gross` of the top 30 production companies using a **Treemap chart**
- Chart added to a worksheet titled **treemap**
