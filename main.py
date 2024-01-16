import pandas as pd
import matplotlib.pyplot as plt
import os
import numpy as np
import textwrap


def number_to_excel_column(column_number):
    result = ""
    while column_number > 0:
        remainder = (column_number - 1) % 26
        result = chr(ord('A') + remainder) + result
        column_number = (column_number - 1) // 26

    return result


def format_tick_label(label):
    if len(label) > 10:
        index = label.rfind(' ', 0, 10)
        if index != -1:
            label = label[:index] + '\n' + label[index+1:]
    return label


output_dir = "output_images"
if os.path.exists(output_dir):
    for file in os.listdir(output_dir):
        file_path = os.path.join(output_dir, file)
        if os.path.isfile(file_path):
            os.remove(file_path)

df = pd.read_excel("Copy of Student Sleep Survey(92).xlsx")

if not os.path.exists(output_dir):
    os.makedirs(output_dir)

histogram_list = []

for colindex, col in enumerate(df.columns[6:]):

    col_data = df[col].dropna()
    col_img_name = col.replace("/", "-")

    if colindex == 11:
        # IMPLEMEMNT THJISSSS
        filtered_data = col_data[col_data.isin(
            ["melatonin", "chinese medicine", "anxiety prescription medicine"])]

        if not filtered_data.empty:
            plt.figure(figsize=(8, 6))
            filtered_data.value_counts().plot.bar()
            plt.title(col)
            plt.xlabel(col)
            plt.ylabel('Frequency')
            plt.xticks(rotation=0)
            plt.savefig(
                f"{output_dir}/{colindex}_{number_to_excel_column(colindex+6)}_{col_img_name}_bar_chart.png")
            plt.close()

    elif col_img_name in ['How would you compare the quality of your sleep during the school year (the past 2 weeks) to your sleep on average outside of the school year', 'Do you have a sleep disorder diagnosed by a medical professional?', 'Do you take any sleep aids-medication?', 'Which (if any) of the following sleep disorders have you suspect you have?'] or (len(col_data.values) < 0.5 * len(df) and colindex != 8):

        fig, ax = plt.subplots(figsize=(8, 14))

        col_data_without_nan = col_data.dropna()

        col_data_without_nan.value_counts().plot.bar()
        middle_index = len(col) // 2
        while middle_index < len(col) and col[middle_index] != ' ':
            middle_index += 1

        if middle_index < len(col):
            col_title = col[:middle_index] + "\n" + col[middle_index+1:]

        max_label_length = max(len(str(label)) for label in df[col].unique())
        wrap_width = 10
        wrapped_labels = [textwrap.fill(str(label), wrap_width)
                          for label in df[col].unique()]
        ax.set_xticks(range(len(df[col].unique())))
        ax.set_xticklabels(wrapped_labels, rotation=0)

        y_max = col_data.value_counts().max()

        if y_max >= 20:
            y_increment = 5
        else:
            y_increment = 2

        y_ticks = range(0, y_max + 1, y_increment)
        ax.set_yticks(y_ticks)

        plt.title(col_title)
        plt.xlabel(col_title)
        plt.ylabel('Frequency')
        plt.xticks(rotation=0)

        plt.savefig(
            f"{output_dir}/{colindex}_{number_to_excel_column(colindex+6)}_{col_img_name}_bar_chart.png")

        plt.close()

    elif col_img_name == 'In the past 2 weeks, how long on average did you sleep every night?':
        # IMPLEMENT
        sleep_durations = []
        for value in col_data:
            lower_bound = int(value.split('-')[0])
            sleep_durations.append(lower_bound)

        plt.figure(figsize=(8, 6))
        plt.hist(sleep_durations, bins=range(0, 10), edgecolor='k', alpha=0.7)

        plt.xlabel('Sleep Duration (hours)')
        plt.ylabel('Frequency')
        plt.title('Sleep Duration Histogram')

        plt.xticks(range(0, 10))

        plt.savefig(
            f"{output_dir}/{colindex}_{number_to_excel_column(colindex+6)}_{col_img_name}_histogram.png")
        histogram_list.append(
            f"{output_dir}/{colindex}_{number_to_excel_column(colindex+6)}_{col_img_name}_histogram.png")

        plt.close()

    elif set(col_data.unique()).issubset({'Yes', 'No'}):
        col_data.value_counts().plot.pie(autopct='%1.1f%%')
        middle_index = len(col) // 2
        while middle_index < len(col) and col[middle_index] != ' ':
            middle_index += 1

        if middle_index < len(col):
            col_title = col[:middle_index] + "\n" + col[middle_index+1:]
        plt.title(col_title)
        plt.ylabel('')
        plt.savefig(
            f"{output_dir}/{colindex}_{number_to_excel_column(colindex+6)}_{col}_pie_chart.png")
        plt.close()

    elif "Agree" in col_data.values or "Somewhat likely" in col_data.values:
        order = []

        unique, counts = np.unique(col_data, return_counts=True)
        dic = dict(zip(unique, counts))
        # {'Agree': 23, 'Disagree': 21, 'Neutral': 5, 'Strongly Agree': 37, 'Strongly Disagree': 6}
        print(dic)

        if "Agree" in col_data.values:
            order = ['Strongly Disagree', 'Disagree',
                     'Neutral', 'Agree', 'Strongly Agree']
        else:
            order = ['Very unlikely', 'Somewhat unlikely',
                     'Neither likely nor unlikely', 'Somewhat likely', 'Very likely']

        y_axis = [col]
        x_order = [dic[i] if i in dic else 0 for i in order]
        print(x_order)  # [6, 21, 5, 23, 37]
        print()

        fig, ax = plt.subplots(figsize=(10, 6))  #

        neutral_position = x_order[3] + x_order[4]
        disagree_position = neutral_position + x_order[2]

        b1 = ax.barh(y_axis, -x_order[0], left=-x_order[2]/2,
                     height=0.3, color=(239/255, 134/255, 54/255))
        b2 = ax.barh(y_axis, -x_order[1], left=-x_order[0]-x_order[2] /
                     2, height=0.3, color=(244/255, 184/255, 121/255))

        b3 = ax.barh(y_axis, x_order[2], left=-x_order[2]/2,
                     height=0.3, color=(229/255, 229/255, 229/255))

        b4 = ax.barh(y_axis, x_order[3], left=x_order[2]/2,
                     height=0.3, color=(132/255, 172/255, 206/255))
        b5 = ax.barh(y_axis, x_order[4], left=x_order[2]/2 +
                     x_order[3], height=0.3, color=(59/255, 117/255, 175/255))

        plt.legend([b1, b2, b3, b4, b5], order, loc="upper right")

        plt.xlabel('Frequency')
        col_title = col
        if "negatively" in col_title and "stress significantly impacts" not in col_title:
            col_title = "School work and stress significantly impacts my " + col_title
        plt.title(col_title)
        plt.savefig(f"{output_dir}/{colindex}_{number_to_excel_column(colindex+6)}_{col_title}_stacked_barh_chart.png",
                    bbox_inches='tight')
        plt.close()

    elif col_data.dtype in ['float64', 'int64']:
        fig, ax = plt.subplots(figsize=(6, 12))
        bin_size = 1
        plt.hist(col_data, bins=int((col_data.max() - col_data.min()) / bin_size), range=(
            col_data.min(), col_data.max() + bin_size), align='left', edgecolor='k', alpha=0.7)
        middle_index = len(col) // 2
        while middle_index < len(col) and col[middle_index] != ' ':
            middle_index += 1

        if middle_index < len(col):
            col = col[:middle_index] + "\n" + col[middle_index+1:]
        plt.title(col)
        plt.xlabel(col)
        plt.ylabel('Frequency')
        plt.savefig(
            f"{output_dir}/{colindex}_{number_to_excel_column(colindex+6)}_{col_img_name}_histogram.png")
        plt.close()

    else:
        print(col_img_name)
        print()

combined_histogram = []
for img_path in histogram_list:
    img = plt.imread(img_path)
    combined_histogram.append(img)


height = sum(img.shape[0] for img in combined_histogram)
width = max(img.shape[1] for img in combined_histogram)

combined_img = np.zeros((height, width, 3), dtype=np.uint8)

y_offset = 0
for img in combined_histogram:
    combined_img[y_offset:y_offset + img.shape[0], :img.shape[1], :] = img
    y_offset += img.shape[0]

combined_histogram_path = os.path.join(output_dir, "combined_histogram.png")
plt.imsave(combined_histogram_path, combined_img)

plt.close("all")
