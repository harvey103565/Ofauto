# Ofauto
Office automation kits based on pywin32.

# Usage
```
from xlneuro import App

app = App()
sheet = app.Worksheets['SHEET NAME']
work_range = sheet.Range('A1:B3')
work_range.Value = ((1, 2), (3, 4), (5, 6))
```
