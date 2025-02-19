# This script generates indivudal reports that summarize data collection since the beginning of the study. It includes ALL time points

# update v2: displays phq9 and gad7 instead of activity and sleep efficiency

import pandas as pd
import os
import matplotlib.pyplot as plt
import numpy as np
import matplotlib.gridspec as gridspec
import matplotlib.dates as mdates
import datetime as dtm
from datetime import datetime, timedelta
from matplotlib.lines import Line2D
import statistics
from matplotlib import rcParams

pd.set_option('display.max_rows', None)

# DEFINE ZONES CUT-OFFS FOR REDCAP QUESTIONNAIRES
phq9_zones = {'Minimal': (0, 5), 'Mild': (5, 10), 'Moderate': (10, 15), 'Moderately Severe': (15, 20), 'Severe': (20, 27)}
phq9_colors = ['green', 'yellow', 'orange', 'red', 'darkred']
gad7_zones = {'Minimal': (0, 5), 'Mild': (5, 11), 'Moderate': (11, 15), 'Severe': (15, 21)}
gad7_colors = ['green', 'yellow', 'orange', 'red']

# Set the font globally
# rcParams['font.family'] = 'Roboto'

# LOAD DATA

# Define a function to add a separator in the legend
def add_legend_separator():
    handles.append(Line2D([0], [0], color='white', lw=0))  # Adds a blank line as separator
    labels.append('')

# Define the path to the folder containing the CSV files
sleep_data_folder_path = r'C:\Users\dylan\OneDrive - McGill University\McPsyt Lab Douglas\actigraphy\sleep_data'
activity_data_folder_path = r'C:\Users\dylan\OneDrive - McGill University\McPsyt Lab Douglas\actigraphy\activity_data'
redcap_data_folder_path = r'C:\Users\dylan\OneDrive - McGill University\McPsyt Lab Douglas\actigraphy\redcap_data'


# List all CSV files in the folder
sleep_files = [file for file in os.listdir(sleep_data_folder_path) if file.endswith('.csv')]
activity_files = [file for file in os.listdir(activity_data_folder_path) if file.endswith('.csv')]
redcap_files = [file for file in os.listdir(redcap_data_folder_path) if file.endswith('.xlsx')]


sleep_data = {}
activity_data = {}
redcap_data = {}
list_of_IDs = []

# Transfer files into dictionnary and create a list of IDs
for sleep_file in sleep_files:
    sleep_file_name = os.path.splitext(sleep_file)[0]
    sleep_data[sleep_file_name] = pd.read_csv(os.path.join(sleep_data_folder_path, sleep_file))

for activity_file in activity_files:
    activity_file_name = os.path.splitext(activity_file)[0]
    activity_data[activity_file_name] = pd.read_csv(os.path.join(activity_data_folder_path, activity_file))

for redcap_file in redcap_files:
    redcap_file_name = os.path.splitext(redcap_file)[0]
    if redcap_file_name[-3:] not in list_of_IDs:
        list_of_IDs.append(redcap_file_name[-3:])
    redcap_data[redcap_file_name] = pd.read_excel(os.path.join(redcap_data_folder_path, redcap_file))


for participant_id in list_of_IDs:

    print(f'PROCESSING PARTICIPANT DD_{participant_id}')
    mean_onset_list = []
    mean_rise_list = []
    sleep_individual_dataframes = []
    sleep_individual_dataframes_names = []
    activity_individual_dataframes = []
    activity_individual_dataframes_names = []

    # Add DataFrames from sleep_data
    for sleep_file_name, df in sleep_data.items():
        if f'{participant_id}' in sleep_file_name:
            sleep_individual_dataframes.append(df)
            sleep_individual_dataframes_names.append(sleep_file_name)
        
    # Add DataFrames from activity_data
    for activity_file_name, df in activity_data.items():
        if f'{participant_id}' in activity_file_name:
            activity_individual_dataframes.append(df)
            activity_individual_dataframes_names.append(activity_file_name)

    if sleep_individual_dataframes and activity_individual_dataframes:
        
        sleep_df = sleep_individual_dataframes[0]
        sleep_df['Sleep.Onset.Time'] = pd.to_datetime(sleep_df['Sleep.Onset.Time'], format='%H:%M', errors='coerce')
        sleep_df['Rise.Time'] = pd.to_datetime(sleep_df['Rise.Time'], format='%H:%M', errors='coerce')
        mean_onset = sleep_df.loc[sleep_df['Night.Starting'] == 'Mean', 'Sleep.Onset.Time'].iloc[0]
        mean_rise = sleep_df.loc[sleep_df['Night.Starting'] == 'Mean', 'Rise.Time'].iloc[0]
        mean_onset_list.append(mean_onset)
        mean_rise_list.append(mean_rise)
        sleep_df = sleep_df[sleep_df['Night.Starting'] != 'Mean']
        sleep_df = sleep_df[(sleep_df == 0).sum(axis=1) < 5]
        sleep_df['Night.Starting'] = pd.to_datetime(sleep_df['Night.Starting'])

        activity_df = activity_individual_dataframes[0]
        # Create a Date column for activity_df based on the dates in the corresponding sleep_df
        activity_df['Date'] = pd.concat([pd.Series(sleep_df['Night.Starting'].iloc[0] - pd.Timedelta(days=1)), sleep_df['Night.Starting']]).reset_index(drop=True)
        activity_df['Date'] = pd.to_datetime(activity_df['Date'])
        activity_df = activity_df[(activity_df==0).sum(axis=1) < 5]

        for i in range(1, len(sleep_individual_dataframes)):
            sleep_individual_dataframes[i]['Sleep.Onset.Time'] = pd.to_datetime(sleep_individual_dataframes[i]['Sleep.Onset.Time'], format='%H:%M', errors='coerce')
            sleep_individual_dataframes[i]['Rise.Time'] = pd.to_datetime(sleep_individual_dataframes[i]['Rise.Time'], format='%H:%M', errors='coerce')
            mean_onset = sleep_individual_dataframes[i].loc[sleep_individual_dataframes[i]['Night.Starting'] == 'Mean', 'Sleep.Onset.Time'].iloc[0]
            mean_rise = sleep_individual_dataframes[i].loc[sleep_individual_dataframes[i]['Night.Starting'] == 'Mean', 'Rise.Time'].iloc[0]
            mean_onset_list.append(mean_onset)
            mean_rise_list.append(mean_rise)
            sleep_individual_dataframes[i] = sleep_individual_dataframes[i][sleep_individual_dataframes[i]['Night.Starting'] != 'Mean']
            sleep_individual_dataframes[i] = sleep_individual_dataframes[i][(sleep_individual_dataframes[i] == 0).sum(axis=1) < 5]

            if len(sleep_individual_dataframes) > 1:
                sleep_df = pd.concat([sleep_df, sleep_individual_dataframes[i]], ignore_index=True)
            
        for i in range(1, len(activity_individual_dataframes)):

            if len(activity_individual_dataframes) > 1:
                # Create a Date column for activity_df based on the dates in the corresponding sleep_df
                sleep_individual_dataframes[i]['Night.Starting'] = pd.to_datetime(sleep_individual_dataframes[i]['Night.Starting'])
                activity_individual_dataframes[i]['Date'] = pd.concat([pd.Series(sleep_individual_dataframes[i]['Night.Starting'].iloc[0] - pd.Timedelta(days=1)), sleep_individual_dataframes[i]['Night.Starting']]).reset_index(drop=True)
                activity_individual_dataframes[i]['Date'] = pd.to_datetime( activity_individual_dataframes[i]['Date'])

                activity_individual_dataframes[i] = activity_individual_dataframes[i][(activity_individual_dataframes[i] == 0).sum(axis=1) < 5]
                activity_df = pd.concat([activity_df, activity_individual_dataframes[i]], ignore_index=True)
                
                # Make sure no dates appear twice, and if so take the sum of the columns
                activity_df = activity_df.groupby('Date', as_index=False)[['Steps', 'Non_Wear', 'Sleep', 'Sedentary', 'Light', 'Moderate', 'Vigorous']].sum()


        
        for redcap_file_name, df in redcap_data.items():
          if redcap_file_name.endswith(participant_id):
            redcap_df = df
            redcap_df_name = redcap_file_name

            # Rename some columns
            redcap_df = redcap_df.rename(columns={
                'Date WHODAS': 'date_whodas',
                'WHODAS': 'whodas',
                'DAST10': 'dast10',
                'Date DAST10': 'date_dast10',
                'AUDIT': 'audit',
                'Date AUDIT': 'date_audit'
            })




        # We will calculate time onset in seconds as the time in seconds after midnight of the day lived before the night of sleep
        # e.g.: if participant sleeps at 20:00, the onset time will be 20*3600 = 72000 seconds
        # e.g.: if participant sleeps at 02:00 AM, the onset time will be (2*3600) + (24*3600) = 93600 seconds
        # We will then calculate the mean of all these values and convert it back to a timestamp
        # For the timestamp, if mean > 86400s (1 day), we will subtract 24 hours to get the correct time
        onset_seconds = []

        for onset in mean_onset_list:
            if onset.hour * 3600 + onset.minute * 60 + onset.second < 43200: # If time is before noon (43200s post-midnight), we add 24 hours
                onset_seconds.append(onset.hour * 3600 + onset.minute * 60 + onset.second + 86400)
            else:
                onset_seconds.append(onset.hour * 3600 + onset.minute * 60 + onset.second)

        mean_onset_seconds = sum(onset_seconds) / len(onset_seconds)

        # If mean_onset_seconds > 86400, subtract 86400 to get the correct time
        if mean_onset_seconds > 86400:
            mean_onset_seconds -= 86400
        final_mean_onset = pd.Timestamp("2023-10-25").replace(second=0) + pd.to_timedelta(mean_onset_seconds, unit='s')

        # Calculate the same for rise time. This is easier, because we know that the rise time is always after midnight.
        rise_seconds = [rise.hour * 3600 + rise.minute * 60 + rise.second for rise in mean_rise_list]
        mean_rise_seconds = sum(rise_seconds) / len(rise_seconds)
        final_mean_rise = pd.Timestamp("2023-10-25").replace(second=0) + pd.to_timedelta(mean_rise_seconds, unit='s')

        print(f'Participant: {participant_id}')
        print(f'Mean onset list: {mean_onset_list}')
        print(f'Final mean onset: {final_mean_onset}')
        print(f'Mean rise list: {mean_rise_list}')
        print(f'final mean rise: {final_mean_rise}')
        

        # Set a specific date (12-10-2003) for both times
        arbitrary_date = datetime.strptime('12-10-2003', '%d-%m-%Y')
        sleep_df['Sleep.Onset.Time'] = sleep_df['Sleep.Onset.Time'].apply(lambda x: x.replace(year=arbitrary_date.year, month=arbitrary_date.month, day=arbitrary_date.day) if pd.notnull(x) else x)
        sleep_df['Rise.Time'] = sleep_df['Rise.Time'].apply(lambda x: x.replace(year=arbitrary_date.year, month=arbitrary_date.month, day=arbitrary_date.day) if pd.notnull(x) else x)

        def adjust_sleep_time(row, date_comparator):
            time = row.time()
            if time < pd.Timestamp('06:00').time() and date_comparator.date() == pd.Timestamp('2003-10-12').date():
                return (date_comparator + pd.Timedelta(days=1)).replace(hour=time.hour, minute=time.minute)
            
            return date_comparator.replace(hour=time.hour, minute=time.minute)

        date_comparator = pd.Timestamp('2003-10-12 06:00')

        # Apply the adjustment to Sleep Time Onset and Rise Time columns
        sleep_df['Sleep.Onset.Time'] = sleep_df['Sleep.Onset.Time'].apply(lambda x: adjust_sleep_time(x, date_comparator))
        sleep_df['Rise.Time'] = sleep_df['Rise.Time'].apply(lambda x: adjust_sleep_time(x, date_comparator))

        onset_variability = sleep_df['Sleep.Onset.Time'].diff().abs().dt.total_seconds().mean() / 3600  # Convert to hours
        rise_variability = sleep_df['Rise.Time'].diff().abs().dt.total_seconds().mean() / 3600  # Convert to hours
        print(onset_variability)
        print(rise_variability)


        # Remove rows with earliest and latest date, as they are incomplete
        # For earliest date, they picked up the watch during the day, so it's not a full day of data
        # For latest date, they are giving back the watch, so it's not a full day of data
        earliest_date = activity_df['Date'].min()
        latest_date = activity_df['Date'].max()
        print(f'The date {earliest_date} has been removed for participant DD_{participant_id}')
        print(f'The date {latest_date} has been removed for participant DD_{participant_id}')
        activity_df = activity_df[activity_df['Date'] != earliest_date]
        activity_df = activity_df[activity_df['Date'] != latest_date]

        # Convert into datetimes again (i know we did it earlier but it seems that it's still strings, so let's just do it again on the contatenated dataframes)
        sleep_df['Night.Starting'] = pd.to_datetime(sleep_df['Night.Starting'])
        activity_df['Date'] = pd.to_datetime(activity_df['Date'])
        redcap_df['date_phq9'] = pd.to_datetime(redcap_df['date_phq9'])
        redcap_df['date_gad7'] = pd.to_datetime(redcap_df['date_gad7'])
        redcap_df['date_cgi-s'] = pd.to_datetime(redcap_df['date_cgi-s'])
        redcap_df['date_saps-sans'] = pd.to_datetime(redcap_df['date_saps-sans'])
        redcap_df['date_whodas'] = pd.to_datetime(redcap_df['date_whodas'])
        redcap_df['date_dast10'] = pd.to_datetime(redcap_df['date_dast10'])
        redcap_df['date_audit'] = pd.to_datetime(redcap_df['date_audit'])


        # Initialize lists to hold valid min and max dates
        min_dates = [sleep_df['Night.Starting'].min(), activity_df['Date'].min(), redcap_df['date_phq9'].min(), redcap_df['date_gad7'].min()]
        max_dates = [sleep_df['Night.Starting'].max(), activity_df['Date'].max(), redcap_df['date_phq9'].max(), redcap_df['date_gad7'].max()]

        # Determine overall min and max date
        min_date = (min(min_dates))
        max_date = (max(max_dates))


        print(f'Min date: {min_date}')
        print(f'Max date: {max_date}')


        # Convert sleep time to hours
        sleep_df['Total.Sleep.Time'] = sleep_df['Total.Sleep.Time'] / 3600

        # Calculate sleep duration variability
        sleep_df['Sleep.Difference'] = sleep_df['Total.Sleep.Time'].diff()
        sleep_variability = sleep_df['Sleep.Difference'].abs().mean()

        # Create a date range for x-ticks from min_date - 1 day to max_date + 1 day
        daily_date_range = pd.date_range(start=min_date - pd.Timedelta(days=1), end=max_date + pd.Timedelta(days=1), freq='D')
        weekly_date_range = pd.date_range(start=min_date - pd.Timedelta(days=1), end=max_date + pd.Timedelta(days=1), freq='W')

        # Define x-limits
        x_limits = (min_date - pd.Timedelta(days=1), max_date + pd.Timedelta(days=1))



        #### Functions to convert sleep onset and rise time in "time after midnight"

        # Function to calculate time difference to midnight
        def calculate_sleep_onset_midnight(time):
            # Convert sleep_onset to datetime (combine with an arbitrary date)
            arbitrary_date = datetime(2000, 1, 1)  # Arbitrary reference date
            time_dt = datetime.combine(arbitrary_date, time)

            # Define midnight
            midnight_dt = datetime.combine(arbitrary_date, datetime.min.time())

            # If sleep onset is between 00:00 and noon
            if time_dt.hour < 12:
                return (time_dt - midnight_dt).total_seconds() / 60

            # If sleep onset is between noon and 23:59:59
            else:
                next_midnight = midnight_dt + timedelta(days=1)
                return -(next_midnight - time_dt).total_seconds() / 60
            
        def calculate_rise_time_midnight(time):
            # Convert sleep_onset to datetime (combine with an arbitrary date)
            arbitrary_date = datetime(2000, 1, 1)  # Arbitrary reference date
            time_dt = datetime.combine(arbitrary_date, time)

            # Define midnight
            midnight_dt = datetime.combine(arbitrary_date, datetime.min.time())

            return (time_dt - midnight_dt).total_seconds() / 60
        
        # Convert time columns to datetime objects
        sleep_df['Sleep.Onset.Time'] = pd.to_datetime(sleep_df['Sleep.Onset.Time'], format='%H:%M', errors='coerce').dt.time
        sleep_df['Rise.Time'] = pd.to_datetime(sleep_df['Rise.Time'], format='%H:%M', errors='coerce').dt.time

        # Apply the function to the DataFrame column
        sleep_df['sleep_onset_minutes'] = sleep_df['Sleep.Onset.Time'].apply(lambda x: calculate_sleep_onset_midnight(x))
        sleep_df['rise_time_minutes'] = sleep_df['Rise.Time'].apply(lambda x: calculate_rise_time_midnight(x))

        mean_sleep_onset_weekdays = sleep_df[sleep_df["Night.Starting"].dt.weekday < 5]['sleep_onset_minutes'].mean()
        mean_sleep_onset_weekends = sleep_df[sleep_df["Night.Starting"].dt.weekday >= 5]['sleep_onset_minutes'].mean()

        mean_rise_time_weekdays = sleep_df[sleep_df["Night.Starting"].dt.weekday < 5]['rise_time_minutes'].mean()
        mean_rise_time_weekends = sleep_df[sleep_df["Night.Starting"].dt.weekday >= 5]['rise_time_minutes'].mean()

        def minutes_to_time(minutes):
          # If minutes are negative, count backwards
          if minutes < 0:
              minutes = 1440 + minutes  # Add 1440 (total minutes in a day) to loop back

          # Convert minutes to hours and minutes
          time = timedelta(minutes=minutes)
          
          # Get the hours and minutes
          hours, remainder = divmod(time.seconds, 3600)
          minutes = remainder // 60

          # Return the time formatted as HH:MM
          return f"{hours:02}:{minutes:02}"

        # Convert the results into time format
        mean_sleep_onset_weekdays = minutes_to_time(mean_sleep_onset_weekdays)
        mean_sleep_onset_weekends = minutes_to_time(mean_sleep_onset_weekends)

        mean_rise_time_weekdays = minutes_to_time(mean_rise_time_weekdays)
        mean_rise_time_weekends = minutes_to_time(mean_rise_time_weekends)










########################### PLOTTING


        # Adjusted Code for Consolidated Legend with Sections
        fig, axes = plt.subplots(2, 2, figsize=(24, 14), sharex=True)

        # Define lists to store legend handles and labels
        handles, labels = [], []

        # Total Sleep Time
        markers = ['x' if d.weekday() >= 6 else 'o' for d in sleep_df['Night.Starting']]

        # Plot all points with a continuous line
        axes[0, 0].plot(sleep_df['Night.Starting'], sleep_df['Total.Sleep.Time'], color='#5D3A9B', label='PHQ-9 Scores (x: Sunday)')

        # Overlay markers for weekdays and weekends
        for date, score, marker in zip(sleep_df['Night.Starting'], sleep_df['Total.Sleep.Time'], markers):
           axes[0, 0].scatter(date, score, marker=marker, color='#5D3A9B')


        avg_sleep = sleep_df['Total.Sleep.Time'].mean()
        avg_sleep_weekdays = sleep_df[sleep_df["Night.Starting"].dt.weekday < 5]["Total.Sleep.Time"].mean()
        avg_sleep_weekends = sleep_df[sleep_df["Night.Starting"].dt.weekday >= 5]["Total.Sleep.Time"].mean()

        axes[0, 0].axhline(avg_sleep, linestyle='-', color='#5D3A9B', label=f'Avg Sleep Duration (total): {avg_sleep:.2f} hrs')
        axes[0, 0].axhline(avg_sleep_weekdays, linestyle='--', color='#5D3A9B', label=f'Avg Sleep Duration (weekdays): {avg_sleep_weekdays:.2f} hrs')
        axes[0, 0].axhline(avg_sleep_weekends, linestyle=':', color='#5D3A9B', label=f'Avg Sleep Duration (weekends): {avg_sleep_weekends:.2f} hrs')

        axes[0, 0].set_ylabel('Time (Hours)')
        axes[0, 0].set_title('Sleep', fontsize=16, fontweight='bold')

        # Collect handles and labels
        h, l = axes[0, 0].get_legend_handles_labels()
        handles.extend(h)
        labels.extend(l)

        # Add separator
        add_legend_separator()

        # PHQ-9 Plot
        phq9_df = redcap_df[['date_phq9', 'phq_9']].dropna()
        for idx, (zone, limits) in enumerate(phq9_zones.items()):
            axes[0, 1].axhspan(limits[0], limits[1], color=phq9_colors[idx], alpha=0.3, label=zone)


        axes[0, 1].plot(phq9_df['date_phq9'], phq9_df['phq_9'], color='blue', label='PHQ-9 Scores')


        mean_phq9 = phq9_df['phq_9'].mean()
        axes[0, 1].axhline(mean_phq9, linestyle='--', color='blue', label=f'Avg PHQ-9 Score: {mean_phq9:.2f}')

        # Add annotations for PHQ-9 values within the filtered data
        for i, score in enumerate(phq9_df['phq_9']):
            if pd.notna(score):
                axes[0, 1].annotate(f'{score:.0f}', 
                            (phq9_df['date_phq9'].iloc[i], score), 
                            textcoords="offset points", 
                            xytext=(0, 10), ha='center')


        axes[0, 1].set_ylabel('PHQ-9 Score')
        axes[0, 1].set_xticks(ticks=daily_date_range, labels=[d.strftime('%Y-%m-%d') if d in weekly_date_range else '' for d in daily_date_range],rotation=90)
        axes[0, 1].set_yticks(np.arange(0, 28, 3))
        axes[0, 1].set_xlim(x_limits)
        axes[0, 1].set_ylim(0, 27)
        axes[0, 1].set_title('PHQ-9 Scores', fontsize=16, fontweight='bold')

        # Custom legend title with detailed info
        title_text = (f'Participant DD_{participant_id}\n\n'
                      'Note: Days of weekend are Saturday and Sunday,\n while nights are Friday nights (leading to Saturday morning)\n and Saturday nights (leading to Sunday morning)\n\n'
            f'Daily Sleep Duration Variability: {sleep_variability:.2f} hrs\n\n'
            f'Mean Sleep Time Onset (total): {final_mean_onset.strftime("%H:%M")}\n'
            f'Mean Sleep Time Onset (weekdays): {mean_sleep_onset_weekdays}\n'
            f'Mean Sleep Time Onset (weekends): {mean_sleep_onset_weekends}\n'
            f'Daily Variability Sleep Time Onset: {onset_variability:.2f} hrs\n\n'
            f'Mean Rise Time (total): {final_mean_rise.strftime("%H:%M")}\n'
            f'Mean Rise Time (weekdays): {mean_rise_time_weekdays}\n'
            f'Mean Rise Time (weekends): {mean_rise_time_weekends}\n'
            f'Daily Variability Rise Time: {rise_variability:.2f} hrs\n')

        # Collect handles and labels
        h, l = axes[0, 1].get_legend_handles_labels()
        handles.extend(h)
        labels.extend(l)

        # Add separator
        add_legend_separator()

        # Steps Plot
        
        markers = ['x' if d.weekday() >= 6 else 'o' for d in activity_df['Date']]

        # Plot all points with a continuous line
        axes[1, 0].plot(activity_df['Date'], activity_df['Steps'], color='#B61826', label='Number of steps (x: Sunday)')

        # Overlay markers for weekdays and weekends
        for date, score, marker in zip(activity_df['Date'], activity_df['Steps'], markers):
           axes[1, 0].scatter(date, score, marker=marker, color='#B61826')


        avg_steps = activity_df['Steps'].mean()
        avg_steps_weekdays = activity_df[activity_df["Date"].dt.weekday < 5]["Steps"].mean()
        avg_steps_weekends = activity_df[activity_df["Date"].dt.weekday >= 5]["Steps"].mean()

        axes[1, 0].axhline(avg_steps, linestyle='-', color='#B61826', label=f'Avg Steps (total): {avg_steps:.0f}')
        axes[1, 0].axhline(avg_steps_weekdays, linestyle='--', color='#B61826', label=f'Avg Steps (weekdays): {avg_steps_weekdays:.0f}')
        axes[1, 0].axhline(avg_steps_weekends, linestyle=':', color='#B61826', label=f'Avg Steps (weekends): {avg_steps_weekends:.0f}')
        
        axes[1, 0].set_ylabel('Steps')
        axes[1, 0].set_title('Steps', fontsize=16, fontweight='bold')



        plt.xlim(x_limits)
        axes[1, 0].set_xticks(ticks=daily_date_range, labels=[d.strftime('%Y-%m-%d') if d in weekly_date_range else '' for d in daily_date_range], rotation=90)
        axes[1, 1].set_xticks(ticks=daily_date_range, labels=[d.strftime('%Y-%m-%d') if d in weekly_date_range else '' for d in daily_date_range], rotation=90)
        axes[1, 0].set_xlabel('Date')
        axes[1, 1].set_xlabel('Date')



        # GAD-7 Plot
        gad7_df = redcap_df[['date_gad7', 'gad_7']].dropna()

        for idx, (zone, limits) in enumerate(gad7_zones.items()):
            axes[1, 1].axhspan(limits[0], limits[1], color=gad7_colors[idx], alpha=0.3)

        # Plot all points with a continuous line
        axes[1, 1].plot(gad7_df['date_gad7'], gad7_df['gad_7'], color='#CB6CE6', label='GAD-7 Score')



        mean_gad7 = gad7_df['gad_7'].mean()
        axes[1, 1].axhline(mean_gad7, linestyle='--', color='#CB6CE6', label=f'Avg GAD-7 Score: {mean_gad7:.2f}')

        # Add annotations for GAD-7 values
        for i, score in enumerate(gad7_df['gad_7']):
            if pd.notna(score):
                axes[1, 1].annotate(f'{score:.0f}', (gad7_df['date_gad7'].iloc[i], score), textcoords="offset points", xytext=(0,10), ha='center')

        axes[1, 1].set_title('GAD-7 Scores', fontsize=16, fontweight='bold')
        axes[1, 1].set_ylabel('GAD-7 Score')
        axes[1, 1].set_ylim(0, 21)
        axes[1, 1].set_xlim(x_limits)
        axes[1, 1].set_yticks(np.arange(0, 19, 3))
        axes[1, 1].set_xticks(ticks=daily_date_range, labels=[d.strftime('%Y-%m-%d') if d in weekly_date_range else '' for d in daily_date_range], rotation=90)

        # Collect handles and labels
        # We'll show GAD first, then steps
        # GAD-7
        h, l = axes[1, 1].get_legend_handles_labels()
        handles.extend(h)
        labels.extend(l)

        # Add separator
        add_legend_separator()

        # Steps
        h, l = axes[1, 0].get_legend_handles_labels()
        handles.extend(h)
        labels.extend(l)




        # Create a single, consolidated legend
        fig.legend(handles, labels, loc='center left', bbox_to_anchor=(0, 0.4), title=title_text, title_fontsize=15, fontsize=15, frameon=True)

        # Adjust the top margin to give space for the title
        plt.subplots_adjust(top=0.75)


        # Align the title text with the legend's bbox
        title_x, title_y = 0, 0.95  # Adjust x and y based on the legend's bbox location
        plt.suptitle('Your Complete Results', 
             fontsize=50, fontweight='bold', ha='left', x=title_x, y=title_y)
        
        
        # Add an underline using fig.text for precise positioning
        fig.text(0, 0.89, '_' * 63, fontsize=50, color='black', ha='left')

        # Add the subtitle below the main title
        fig.text(0, 0.78, 
                'Information from the watch and questionnaires data since the beginning of the study.\nFrom it, we can track sleep variation, number of steps, depressive symptoms severity (PHQ-9), and anxiety symptoms severity (GAD-7)\nOnly data from the past month are included in the current progress report.',
                fontsize=25, ha='left')


        


        # Adjust layout
        plt.subplots_adjust(left=0.3, right=0.95)
        plt.savefig(f'../individual_overall_reports/overall_report_DD_{participant_id}.png', dpi=300)
    














    else:
        print(f'Participant DD_{participant_id} has no actigraphy data.')


        for redcap_file_name, df in redcap_data.items():
          if redcap_file_name.endswith(participant_id):
            redcap_df = df
            redcap_df_name = redcap_file_name

            # Rename some columns
            redcap_df = redcap_df.rename(columns={
                'Date WHODAS': 'date_whodas',
                'WHODAS': 'whodas',
                'DAST10': 'dast10',
                'Date DAST10': 'date_dast10',
                'AUDIT': 'audit',
                'Date AUDIT': 'date_audit'
            })


        # If there is no data for PHQ-9 or GAD-7, skip the participant
        if redcap_df['phq_9'].count() == 0 and redcap_df['gad_7'].count() == 0:
            print(f'Participant DD_{participant_id} has no PHQ-9 or GAD-7 data.')
            continue

        redcap_df['date_phq9'] = pd.to_datetime(redcap_df['date_phq9'])
        redcap_df['date_gad7'] = pd.to_datetime(redcap_df['date_gad7'])

         # Initialize lists to hold valid min and max dates
        min_dates = [redcap_df['date_phq9'].min(), redcap_df['date_gad7'].min()]
        max_dates = [redcap_df['date_phq9'].max(), redcap_df['date_gad7'].max()]

        # Determine overall min and max date
        min_date = (min(min_dates))
        max_date = (max(max_dates))


        print(f'Min date: {min_date}')
        print(f'Max date: {max_date}')

        # Create a date range for x-ticks from min_date - 1 day to max_date + 1 day
        daily_date_range = pd.date_range(start=min_date - pd.Timedelta(days=1), end=max_date + pd.Timedelta(days=1), freq='D')
        weekly_date_range = pd.date_range(start=min_date - pd.Timedelta(days=1), end=max_date + pd.Timedelta(days=1), freq='W')

        # Define x-limits
        x_limits = (min_date - pd.Timedelta(days=1), max_date + pd.Timedelta(days=1))

        ########################### PLOTTING


        # Adjusted Code for Consolidated Legend with Sections
        fig, axes = plt.subplots(2, 2, figsize=(24, 14), sharex=True)

        # Set sleep and steps subplots as empty
        axes[0, 0].set_xticks([])
        axes[0, 0].set_yticks([])
        axes[0, 0].set_ylabel('Time (Hours)')
        axes[0, 0].set_title('Sleep', fontsize=16, fontweight='bold')


        axes[1, 0].set_xticks([])
        axes[1, 0].set_yticks([])
        axes[1, 0].set_ylabel('Steps')
        axes[1, 0].set_title('Steps', fontsize=16, fontweight='bold')
        axes[1, 0].set_xticks(ticks=daily_date_range, labels=[d.strftime('%Y-%m-%d') if d in weekly_date_range else '' for d in daily_date_range], rotation=90)
        axes[1, 0].set_xlabel('Date')



        # Define lists to store legend handles and labels
        handles, labels = [], []


        # PHQ-9 Plot
        phq9_df = redcap_df[['date_phq9', 'phq_9']].dropna()
        for idx, (zone, limits) in enumerate(phq9_zones.items()):
            axes[0, 1].axhspan(limits[0], limits[1], color=phq9_colors[idx], alpha=0.3, label=zone)

        # Plot all points with a continuous line
        axes[0, 1].plot(phq9_df['date_phq9'], phq9_df['phq_9'], color='blue', label='PHQ-9 Scores')

        mean_phq9 = phq9_df['phq_9'].mean()
        axes[0, 1].axhline(mean_phq9, linestyle='--', color='blue', label=f'Avg PHQ-9 Score: {mean_phq9:.2f}')

        # Add annotations for PHQ-9 values within the filtered data
        for i, score in enumerate(phq9_df['phq_9']):
            if pd.notna(score):
                axes[0, 1].annotate(f'{score:.0f}', 
                            (phq9_df['date_phq9'].iloc[i], score), 
                            textcoords="offset points", 
                            xytext=(0, 10), ha='center')


        axes[0, 1].set_ylabel('PHQ-9 Score')
        axes[0, 1].set_xticks(ticks=daily_date_range, labels=[d.strftime('%Y-%m-%d') if d in weekly_date_range else '' for d in daily_date_range],rotation=90)
        axes[0, 1].set_yticks(np.arange(0, 28, 3))
        axes[0, 1].set_xlim(x_limits)
        axes[0, 1].set_ylim(0, 27)
        axes[0, 1].set_title('PHQ-9 Scores', fontsize=16, fontweight='bold')

        # Custom legend title with detailed info
        title_text = (f'Participant DD_{participant_id}')

        # Collect handles and labels
        h, l = axes[0, 1].get_legend_handles_labels()
        handles.extend(h)
        labels.extend(l)

        # Add separator
        add_legend_separator()

        # GAD-7 Plot
        gad7_df = redcap_df[['date_gad7', 'gad_7']].dropna()

        for idx, (zone, limits) in enumerate(gad7_zones.items()):
            axes[1, 1].axhspan(limits[0], limits[1], color=gad7_colors[idx], alpha=0.3)

        # Plot all points with a continuous line
        axes[1, 1].plot(gad7_df['date_gad7'], gad7_df['gad_7'], color='#CB6CE6', label='GAD-7 Score')




        mean_gad7 = gad7_df['gad_7'].mean()
        axes[1, 1].axhline(mean_gad7, linestyle='--', color='#CB6CE6', label=f'Avg GAD-7 Score: {mean_gad7:.2f}')

        # Add annotations for GAD-7 values
        for i, score in enumerate(gad7_df['gad_7']):
            if pd.notna(score):
                axes[1, 1].annotate(f'{score:.0f}', (gad7_df['date_gad7'].iloc[i], score), textcoords="offset points", xytext=(0,10), ha='center')

        axes[1, 1].set_title('GAD-7 Scores', fontsize=16, fontweight='bold')
        axes[1, 1].set_ylabel('GAD-7 Score')
        axes[1, 1].set_ylim(0, 21)
        axes[1, 1].set_xlim(x_limits)
        axes[1, 1].set_yticks(np.arange(0, 19, 3))
        axes[1, 1].set_xticks(ticks=daily_date_range, labels=[d.strftime('%Y-%m-%d') if d in weekly_date_range else '' for d in daily_date_range], rotation=90)

        # Collect handles and labels
        # We'll show GAD first, then steps
        # GAD-7
        h, l = axes[1, 1].get_legend_handles_labels()
        handles.extend(h)
        labels.extend(l)

        

        # Create a single, consolidated legend
        fig.legend(handles, labels, loc='center left', bbox_to_anchor=(0, 0.5), title=title_text, title_fontsize=15, fontsize=15, frameon=True)

        # Adjust the top margin to give space for the title
        plt.subplots_adjust(top=0.75)


        # Align the title text with the legend's bbox
        title_x, title_y = 0, 0.95  # Adjust x and y based on the legend's bbox location
        plt.suptitle('Your Complete Results', 
             fontsize=50, fontweight='bold', ha='left', x=title_x, y=title_y)
        
        
        # Add an underline using fig.text for precise positioning
        fig.text(0, 0.89, '_' * 63, fontsize=50, color='black', ha='left')

        # Add the subtitle below the main title
        fig.text(0, 0.78, 
                'Information from the watch and questionnaires data since the beginning of the study.\nFrom it, we can track sleep variation, number of steps, depressive symptoms severity (PHQ-9), and anxiety symptoms severity (GAD-7)\nOnly data from the past month are included in the current progress report.',
                fontsize=25, ha='left')


        


        # Adjust layout
        plt.subplots_adjust(left=0.25, right=0.95)
        plt.savefig(f'../individual_overall_reports/overall_report_DD_{participant_id}.png', dpi=300)