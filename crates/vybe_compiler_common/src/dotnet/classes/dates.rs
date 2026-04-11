//! Date / time pickers.
//!
//! Both `DateTimePicker` and `MonthCalendar` inherit from `Control`.

use super::DotnetClass;

pub fn classes() -> &'static [DotnetClass] {
    &[
        DotnetClass {
            name: "DateTimePicker",
            parent: Some("Control"),
            properties: &[
                "CalendarFont",
                "CalendarForeColor",
                "CalendarMonthBackground",
                "CalendarTitleBackColor",
                "CalendarTitleForeColor",
                "CalendarTrailingForeColor",
                "Checked",
                "CustomFormat",
                "DropDownAlign",
                "Format",
                "MaxDate",
                "MinDate",
                "PreferredHeight",
                "ShowCheckBox",
                "ShowUpDown",
                "Value",
            ],
            widget_host_fn: Some("new_DateTimePicker"),
        },
        DotnetClass {
            name: "MonthCalendar",
            parent: Some("Control"),
            properties: &[
                "AnnuallyBoldedDates",
                "BoldedDates",
                "CalendarDimensions",
                "FirstDayOfWeek",
                "MaxDate",
                "MaxSelectionCount",
                "MinDate",
                "MonthlyBoldedDates",
                "ScrollChange",
                "SelectionEnd",
                "SelectionRange",
                "SelectionStart",
                "ShowToday",
                "ShowTodayCircle",
                "ShowWeekNumbers",
                "TitleBackColor",
                "TitleForeColor",
                "TodayDate",
                "TodayDateSet",
                "TrailingForeColor",
            ],
            widget_host_fn: Some("new_MonthCalendar"),
        },
    ]
}
