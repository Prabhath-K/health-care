# health-care

# Medical Data Cleaning

##  Overview

This project focuses on cleaning and preparing medical and hospitalization data using **Microsoft Excel**. The dataset contains two main tables:

- **Medical Examinations**
- **Hospitalization Details**

The primary objective is to identify and handle missing values represented by `?`, replace them with appropriate values, and improve the overall quality of the dataset before further analysis.

---

##  Data Cleaning Tasks

### 1. Identify Missing Values

The first step was to check the number of missing values represented by `?` in each column of both the **Medical Examinations** and **Hospitalization Details** tables.

The `COUNTIF` function was used to count missing values.

```excel
=COUNTIF(range,"?")
```

Example:

```excel
=COUNTIF('Hospitalisation Details'!A:A,"?")
```

This formula counts all cells containing `?` in the selected column.

A separate solution table was created to summarize the number of missing values found in each column.

---

### 2. Handle Missing Month and Year Values

Missing values in the `month` and `year` columns were handled separately.

#### Month

Missing values in the `month` column were replaced with:

```text
Sep
```

This was done using Excel's **Find and Replace** feature:

1. Select the `month` column.
2. Press `Ctrl + H`.
3. In **Find what**, enter `?`.
4. In **Replace with**, enter `Sep`.
5. Click **Replace All**.

#### Year

For the `year` column, missing values were replaced with the average year rounded to the nearest integer.

The average was calculated using:

```excel
=ROUND(AVERAGE(B:B),0)
```

After calculating the rounded average, the missing `?` values in the `year` column were replaced with the calculated value.

---

### 3. Fill Missing Values Using the Most Frequent Value

Missing values in the following columns were filled using their **most frequently occurring value (mode)**:

- `smoker`
- `Hospital tier`
- `City tier`

#### Hospital Tier

The frequency of each hospital tier was calculated using `COUNTIF`.

```excel
=COUNTIF('Hospitalisation Details'!G:G,"tier - 1")
```

```excel
=COUNTIF('Hospitalisation Details'!G:G,"tier - 2")
```

```excel
=COUNTIF('Hospitalisation Details'!G:G,"tier - 3")
```

The tier with the highest count was identified as the most frequently occurring value and used to replace missing values.

#### City Tier

The same method was applied to the `City tier` column by calculating the frequency of each tier and selecting the one with the highest occurrence.

#### Smoker

The `smoker` column contains two main values:

```text
yes
no
```

A **Pivot Table** was used to calculate their frequencies.

Steps:

1. Select the dataset.
2. Go to **Insert → PivotTable**.
3. Add `smoker` to the **Rows** field.
4. Add `smoker` again to the **Values** field and use **Count**.
5. Compare the counts of `yes` and `no`.
6. Replace missing values with the most frequently occurring value.

---

### 4. Handle Missing State ID Values

Missing values in the `State ID` column were replaced with:

```text
Unknown
```

This ensures that records with unavailable state information are retained instead of being removed.

Using **Find and Replace**:

1. Select the `State ID` column.
2. Press `Ctrl + H`.
3. In **Find what**, enter:

```text
~?
```

4. In **Replace with**, enter:

```text
Unknown
```

5. Click **Replace All**.

> In Excel, `?` is treated as a wildcard character. Therefore, `~?` is used when searching for the actual question-mark character.

---

##  Tools & Techniques Used

- Microsoft Excel
- `COUNTIF`
- `AVERAGE`
- `ROUND`
- Pivot Tables
- Find and Replace
- Missing-value imputation
- Mode/Frequency analysis

---

##  Data Cleaning Summary

| Column / Task | Missing Value Strategy |
|---|---|
| Missing-value identification | `COUNTIF(range,"?")` |
| Month | Replaced with `Sep` |
| Year | Replaced with rounded average year |
| Smoker | Replaced with most frequent value using Pivot Table |
| Hospital Tier | Replaced with most frequent tier |
| City Tier | Replaced with most frequent tier |
| State ID | Replaced with `Unknown` |

---

##  Outcome

After completing the cleaning process:

- Missing values were identified across both datasets.
- Missing month values were replaced with `Sep`.
- Missing year values were replaced with the rounded average year.
- Missing categorical values were replaced using their most frequently occurring values.
- Missing `State ID` values were categorized as `Unknown`.
- The resulting dataset is cleaner and more suitable for further analysis and visualization.

---

##  Dataset Tables

The analysis primarily uses:

```text
Medical Examinations
Hospitalization Details
```

These cleaned tables can now be used for additional exploratory data analysis, statistical analysis, dashboard creation, and visualization.
