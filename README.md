# Conference Data Management

## Project Overview
This project is designed to streamline attendee management, forecast event attendance and costs, optimize pricing strategies, and create an interactive dashboard for a professional annual conference. The solution includes Excel automation using VBA, forecasting with built-in Excel functions, and an interactive dashboard.

## Features

### Attendee Management (VBA UserForm)
- **UserForm Components:**
  - TextBoxes for attendee details
  - ComboBoxes for session preferences and ticket pricing
  - RadioButtons for online/in-person attendance
  - ListBox for dietary preferences
  - Buttons for Submit, Fetch, Update, and Cancel actions
- **VBA Functionalities:**
  - Submit: Saves attendee details to the dataset.
  - Fetch: Allows editing of existing records.
  - Update: Saves changes to an existing record.
  - Cancel: Closes the form.

### Forecast Attendance and Costs
- Summarized total ticket sales and registrations per month.
- Used `FORECAST.ETS()` to predict future attendance.

### Goal Seek for Pricing Strategies
- Calculated forecasted revenue: `=Average Ticket Price * Predicted Attendance`
- Determined profit.
- Used **Goal Seek** to:
  - Found the ticket price required to achieve the profit target.
  - Determined the total attendees required to meet the target.

### Interactive Dashboard
- **Metrics:** Total attendees and revenue.
- **Charts:**
  - Bar chart: Revenue by session type
  - Line chart: Attendee trends over time
  - Pie chart: Online vs. In-person attendance
- **Filters:**
  - Slicers for attendee type and session preferences.
  - Timeline filter for period-based data visualization.

### Automate Updates with Macros
- Used VBA macro to refresh all PivotTables and charts.
- Assigned a button labeled **Refresh Dashboard** to execute the macro.

## Notes
This project ensures accurate financial forecasting and enhances event planning through automation and data analysis.
