# Import the redcap dataframe and integrate it with longitudinal data

import pandas as pd
import numpy as np
from datetime import datetime
from datetime import datetime, timedelta, time
import os

pd.set_option('display.max_rows', None)


filepath = r'C:\Users\dylan\OneDrive - McGill University\McPsyt Lab Douglas\actigraphy'
redcap_extraction_df = pd.read_csv(os.path.join(filepath, "redcap_raw.csv"))

redcap_extraction_df['participant_id'] = redcap_extraction_df.groupby('record_id')['participant_id'].ffill()

# Make it so that record_id=32 is DD_10_2024_016 (confirmed with Ashley)
redcap_extraction_df.loc[redcap_extraction_df['record_id'] == 32, 'participant_id'] = 'DD_10_2024_016'



# Drop the duplicate phq-9 score of participant DD_005. We drop the earliest one, which is the one with the timestamp 2024-10-31 18:39:15.
redcap_extraction_df = redcap_extraction_df[
    ~((redcap_extraction_df['participant_id'] == 'DD_07_2024_005') & 
      (redcap_extraction_df['phq9_timestamp'] == '2024-10-31 18:39:15') & 
      (redcap_extraction_df['phq9_fs'] == 14.0))
]


redcap_extraction_df = redcap_extraction_df[[
  'participant_id', 'record_id',

  # English columns
  'wsas_fs', 'wsas_timestamp',
  'madrs_fs', 'madrs_timestamp',
  'saps_global', 'saps_fs', 'sans_global', 'sans_fs', 'sapssans_timestamp',
  'saps_7', 'saps_20', 'saps_25', 'saps_34',
  'sans_8', 'sans_13', 'sans_17', 'sans_22', 'sans_25',
  'ymrs_fs', 'ymrs_timestamp',
  'qidssr16_fs', 'qidssr16_timestamp',
  'cssrs_fs', 'cssrs_timestamp',
  'phq9_9', 'phq9_timestamp',
  'phq9_fs',
  'gad7_fs', 'gad7_timestamp',
  'audit_fs', 'audit_timestamp',
  'dast10_fs', 'dast10_timestamp',
  'whodas_fs', 'whodas_timestamp',
  'cgi_s', 'cgi_timestamp',

  # French columns
  'wsas_fs_fr', 'wsasfr_timestamp',
  'madrsfs_fr', 'madrsfr_timestamp',
  'saps_global_fr', 'saps_fs_fr', 'sans_global_fr', 'sans_fs_fr', 'sapssansfr_timestamp',
  'saps_7_fr', 'saps_20_fr', 'saps_25_fr', 'saps_34_fr',
  'sans_8_fr', 'sans_13_fr', 'sans_17_fr', 'sans_22_fr', 'sans_25_fr',
  'ymrsfs_fr', 'ymrsfr_timestamp',
  'qidssr16_fs_fr', 'qidssr16fr_timestamp',
  'cssrs_fs_fr', 'cssrs_fr_timestamp',
  'phq9_fs_fr', 'phq9fr_timestamp',
  'gad7_fs_fr', 'gad7fr_timestamp',
  'audit_fs_fr', 'auditfr_timestamp',
  'dast10_fsfr', 'dast10fr_timestamp',
  'whodas_fs_fr', 'whodasfr_timestamp',
  'cgi_s_fr', 'cgifr_timestamp',
  ]]

redcap_extraction_df.loc[redcap_extraction_df['record_id'] == 32, 'participant_id'] = 'DD_10_2024_016'


redcap_extraction_df.rename(columns={
  # English columns
  'wsas_fs': 'wsas',
  'madrs_fs': 'madrs',
  'wsas_timestamp': 'date_wsas',
  'madrs_timestamp': 'date_madrs',
  'phq9_timestamp': 'date_phq9',
  'saps_global': 'saps_global',
  'saps_fs': 'saps_total',
  'sans_global': 'sans_global',
  'sans_fs': 'sans_total',
  'sapssans_timestamp': 'date_saps-sans',
  'ymrs_fs': 'ymrs',
  'ymrs_timestamp': 'date_ymrs',
  'qidssr16_fs': 'qids',
  'qidssr16_timestamp': 'date_qids',
  'cssrs_fs': 'cssrs',
  'cssrs_timestamp': 'date_cssrs',
  'phq9_fs': 'phq_9',
  'gad7_fs': 'gad_7',
  'gad7_timestamp': 'date_gad7',
  'saps_7': 'saps_hallucinations',
  'saps_20': 'saps_delusions',
  'saps_25': 'saps_bizarre_behavior',
  'saps_34': 'saps_positive_thought_disorder',
  'sans_8': 'sans_flattening',
  'sans_13': 'sans_alogia',
  'sans_17': 'sans_apathy',
  'sans_22': 'sans_asociality',
  'sans_25': 'sans_attention',
  'audit_fs': 'audit',
  'audit_timestamp': 'date_audit',
  'dast10_fs': 'dast10',
  'dast10_timestamp': 'date_dast10',
  'whodas_fs': 'whodas',
  'whodas_timestamp': 'date_whodas',
  'cgi_s': 'cgi_s',
  'cgi_timestamp': 'date_cgis',

  # French columns
  'wsas_fs_fr': 'wsas_fr',
  'madrsfs_fr': 'madrs_fr',
  'wsasfr_timestamp': 'date_wsas_fr',
  'madrsfr_timestamp': 'date_madrs_fr',
  'phq9fr_timestamp': 'date_phq9_fr',
  'saps_global_fr': 'saps_global_fr',
  'saps_fs_fr': 'saps_total_fr',
  'sans_global_fr': 'sans_global_fr',
  'sans_fs_fr': 'sans_total_fr',
  'sapssansfr_timestamp': 'date_saps-sans_fr',
  'ymrsfs_fr': 'ymrs_fr',
  'ymrsfr_timestamp': 'date_ymrs_fr',
  'qidssr16_fs_fr': 'qids_fr',
  'qidssr16fr_timestamp': 'date_qids_fr',
  'cssrs_fs_fr': 'cssrs_fr',
  'cssrs_fr_timestamp': 'date_cssrs_fr',
  'phq9_fs_fr': 'phq_9_fr',
  'gad7_fs_fr': 'gad_7_fr',
  'gad7fr_timestamp': 'date_gad7_fr',
  'audit_fs_fr': 'audit_fr',
  'auditfr_timestamp': 'date_audit_fr',
  'dast10_fsfr': 'dast10_fr',
  'dast10fr_timestamp': 'date_dast10_fr',
  'whodas_fs_fr': 'whodas_fr',
  'whodasfr_timestamp': 'date_whodas_fr',
  'saps_7_fr': 'saps_hallucinations_fr',
  'saps_20_fr': 'saps_delusions_fr',
  'saps_25_fr': 'saps_bizarre_behavior_fr',
  'saps_34_fr': 'saps_positive_thought_disorder_fr',
  'sans_8_fr': 'sans_flattening_fr',
  'sans_13_fr': 'sans_alogia_fr',
  'sans_17_fr': 'sans_apathy_fr',
  'sans_22_fr': 'sans_asociality_fr',
  'sans_25_fr': 'sans_attention_fr',
  'cgi_s_fr': 'cgi_s_fr',
  'cgifr_timestamp': 'date_cgis_fr'
}, inplace=True)

for col in redcap_extraction_df.columns:
  if 'date_' in col:
    redcap_extraction_df[col] = pd.to_datetime(redcap_extraction_df[col], errors='coerce')
    redcap_extraction_df[col] = redcap_extraction_df[col].dt.date

print(redcap_extraction_df.head())

fr_cols = [col for col in redcap_extraction_df.columns if col.endswith('_fr')]
non_fr_cols = [col for col in redcap_extraction_df.columns if not col.endswith('_fr') and col != 'participant_id']



# Let's separate the whole dataframe into one for each participant

for participant in redcap_extraction_df['participant_id'].unique():
  participant_df = redcap_extraction_df[redcap_extraction_df['participant_id'] == participant]
  participant_df = participant_df.reset_index(drop=True)

  participant_df = participant_df.dropna(subset=[col for col in participant_df.columns if col not in ['participant_id', 'record_id']], how='all')

  # Drop record_id column
  participant_df = participant_df.drop(columns=['record_id'])

  # Determine if the participant answered the questionnaire in English or French

  # Check if all '_fr' columns are NaN for 'participant_id'
  if participant_df.loc[participant_df['participant_id'].notna(), fr_cols].isna().all(axis=1).all():
      # Drop all '_fr' columns
      participant_df = participant_df.drop(columns=fr_cols)
  else:
      # Keep only '_fr' columns and 'participant_id'
      participant_df = participant_df[['participant_id'] + fr_cols]

  # Remove '_fr' suffix from columns
  participant_df.columns = [col.replace('_fr', '') for col in participant_df.columns]


  print(participant_df.head())

  # Save the dataframe to a csv file
  participant_df.to_csv(os.path.join(f'{filepath}/redcap_data', f"redcap_data_{participant}.csv"), index=False)