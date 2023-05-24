# begin
import pandas as pd
import numpy as np
from sklearn.linear_model import LinearRegression

output_xls_name = 'output_new_2.xlsx'
input_xls_name = 'akb_v2023.xlsx'


def power4_regression(x, y):
    # Apply logarithmic transformation
    log_x = np.log(x)
    log_y = np.log(y)

    # Fit linear regression model
    model = LinearRegression()
    model.fit(log_x.reshape(-1, 1), log_y)

    # Extract coefficients
    intercept = np.exp(model.intercept_)
    coef = model.coef_[0]

    # Calculate predicted values
    # for test puprouses - AVG REL ERROR
    # y_pred = intercept * np.power(x, coef)

    # Calculate average relative error in percentage
    # for test purpouses
    # relative_error = np.abs((y - y_pred) / y) * 100
    # avg_relative_error = np.mean(relative_error)

    # return intercept, coef, avg_relative_error
    return intercept, coef


def get_cubic_regression_coefficients(x, y):
    # Get the cubic regression coefficients.
    coefficients = np.polyfit(x, y, 3)
    coefficients_long = ['{:.12f}'.format(coefficient) for coefficient in coefficients]
    return coefficients_long


# rename dict for minutes values without hr marks, current values become .1
rename_dict = {
    '5': '5', '10': '10', '15': '15', '30': '30', '45': '45',
    '60/1ч': '60', '120/2ч': '120', '180/3ч': '180', '300/5ч': '300', '480/8ч': '480', '600/10ч': '600',
    '1200/20ч': '1200', '60/1ч.1': '60.1', '120/2ч.1': '120.1', '180/3ч.1': '120.1', '300/5ч.1': '300.1',
    '480/8ч.1': '480.1', '600/10ч.1': '600.1', '1200/20ч.1': '1200.1'
}


xl = pd.ExcelFile(input_xls_name)
df = xl.parse(xl.sheet_names[1])
# remove C-C (russian / english misspelling)
df = df.replace("RС", "RC", regex=True)


# get list of minutes
minutes_uncleared = (df.iloc[2:2, 2:14].columns.tolist())
minutes_list = []
for values in minutes_uncleared:
    minutes_list.append(str(values).split('/')[0])

# change minutes values to floats
for i in range(len(minutes_list)):
    minutes_list[i] = float(minutes_list[i])

# new dataframe for regression coefficients
new_df = pd.DataFrame(columns=['p_45_180_a', 'p_45_180_b', 'p_45_180_c', 'p_45_180_d', 'p_45_180_all',
                               'p_180_1200_a', 'p_180_1200_b', 'p_180_1200_all',
                               'i_45_180_a', 'i_45_180_b', 'i_45_180_c', 'i_45_180_d', 'i_45_180_all',
                               'i_180_1200_a', 'i_180_1200_b', 'i_180_1200_all'])

# iterate over batt P and I values database
for i, row in df.iterrows():
    # get discharge value in the start of the row
    discharge_value = row.values[1]
    if (isinstance(row.values[1], float)) and (1.9 >= discharge_value >= 1.5):

        # power coefficients
        column_values_45_180 = list(row.iloc[6:10])
        minutes_45_180 = minutes_list[4:8]
        p_45_180_a, p_45_180_b, p_45_180_c, p_45_180_d = \
            get_cubic_regression_coefficients(minutes_45_180, column_values_45_180)
        p_45_180_all = ', '.join([str(p_45_180_a), str(p_45_180_b), str(p_45_180_c), str(p_45_180_d)])

        column_values_180_1200 = list(row.iloc[9:14])
        minutes_180_1200 = minutes_list[7:12]
        p_180_1200_a, p_180_1200_b = power4_regression(minutes_180_1200, column_values_180_1200)
        p_180_1200_all = ', '.join([str(p_180_1200_a), str(p_180_1200_a)])

        # current (I) coefficients
        column_values_45_180 = list(row.iloc[23:27])
        minutes_45_180 = minutes_list[4:8]
        i_45_180_a, i_45_180_b, i_45_180_c, i_45_180_d = \
            get_cubic_regression_coefficients(minutes_45_180, column_values_45_180)
        i_45_180_all = ', '.join([str(i_45_180_a), str(i_45_180_b), str(i_45_180_c), str(i_45_180_d)])
        column_values_180_1200 = list(row.iloc[26:31])
        minutes_180_1200 = minutes_list[7:12]
        i_180_1200_a, i_180_1200_b = power4_regression(minutes_180_1200, column_values_180_1200)
        i_180_1200_all = ', '.join([str(i_180_1200_a), str(i_180_1200_a)])

        # store all the coefficients of one row in a new row
        data = {'p_45_180_a': p_45_180_a, 'p_45_180_b': p_45_180_b, 'p_45_180_c': p_45_180_c, 'p_45_180_d': p_45_180_d,
                'p_45_180_all': p_45_180_all,
                'p_180_1200_a': p_180_1200_a, 'p_180_1200_b': p_180_1200_b, 'p_180_1200_all': p_180_1200_all,
                'i_45_180_a': i_45_180_a, 'i_45_180_b': i_45_180_b, 'i_45_180_c': i_45_180_c, 'i_45_180_d': i_45_180_d,
                'i_45_180_all': i_45_180_all,
                'i_180_1200_a': i_180_1200_a, 'i_180_1200_b': i_180_1200_b, 'i_180_1200_all': i_180_1200_all}
        # append new row to a database
        new_df = new_df.append(data, ignore_index=True)

    else:
        data = {
            'p_45_180_a': '', 'p_45_180_b': '', 'p_45_180_c': '', 'p_45_180_d': '', 'p_45_180_all': '',
            'p_180_1200_a': '', 'p_180_1200_b': '', 'p_180_1200_all': '',
            'i_45_180_a': '', 'i_45_180_b': '', 'i_45_180_c': '', 'i_45_180_d': '', 'i_45_180_all': '',
            'i_180_1200_a': '', 'i_180_1200_b': '', 'i_180_1200_all': ''
        }
        # append empty string where is no data
        new_df = new_df.append(data, ignore_index=True)

# save new xlsx file with coefficients
new_df.to_excel(output_xls_name, index=False)
