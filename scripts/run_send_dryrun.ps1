param(
  [Parameter(Mandatory=$true)][string]$To,
  [string]$Csv="contacts.csv",
  [int]$DailyCap=10,
  [int]$RatePerSecond=1
)
$env:CSV_PATH=$Csv
$env:DAILY_CAP="$DailyCap"
$env:RATE_PER_SECOND="$RatePerSecond"
$env:DRY_RUN="1"
$env:DRY_RUN_TO=$To
python .\send.py
