param(
  [string]$Csv="contacts.csv",
  [string]$Template="followup_v1",
  [int]$RatePerSecond=1
)
$env:CSV_PATH=$Csv
$env:FOLLOWUP_TEMPLATE=$Template
$env:RATE_PER_SECOND="$RatePerSecond"
python .\followup.py
