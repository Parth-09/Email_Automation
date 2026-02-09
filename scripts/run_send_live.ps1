param(
  [string]$Csv="contacts.csv",
  [int]$DailyCap=100,
  [int]$RatePerSecond=1
)
$env:CSV_PATH=$Csv
$env:DAILY_CAP="$DailyCap"
$env:RATE_PER_SECOND="$RatePerSecond"
Remove-Item Env:DRY_RUN -ErrorAction SilentlyContinue
Remove-Item Env:DRY_RUN_TO -ErrorAction SilentlyContinue
python .\send.py
