use crate::value::{RuntimeError, Value};
use std::time::SystemTime;
use chrono::{Datelike, Timelike};

fn get_local_time() -> (i32, i32, i32, i32, i32, i32) {
    // Get seconds since epoch and compute date/time components
    let secs = SystemTime::now()
        .duration_since(SystemTime::UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs() as i64;

    // Simple date calculation (no timezone handling, uses UTC)
    let days = secs / 86400;
    let time_of_day = secs % 86400;

    let hour = (time_of_day / 3600) as i32;
    let minute = ((time_of_day % 3600) / 60) as i32;
    let second = (time_of_day % 60) as i32;

    // Days since 1970-01-01
    let mut y = 1970i32;
    let mut remaining_days = days;

    loop {
        let days_in_year = if y % 4 == 0 && (y % 100 != 0 || y % 400 == 0) { 366 } else { 365 };
        if remaining_days < days_in_year {
            break;
        }
        remaining_days -= days_in_year;
        y += 1;
    }

    let leap = y % 4 == 0 && (y % 100 != 0 || y % 400 == 0);
    let month_days = [31, if leap { 29 } else { 28 }, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    let mut m = 0;
    for md in &month_days {
        if remaining_days < *md {
            break;
        }
        remaining_days -= md;
        m += 1;
    }

    (y, m + 1, remaining_days as i32 + 1, hour, minute, second)
}

/// Now() - Returns current date and time as string
pub fn now_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if !args.is_empty() {
        return Err(RuntimeError::Custom("Now takes no arguments".to_string()));
    }
    let (y, m, d, h, min, s) = get_local_time();
    Ok(Value::String(format!("{}/{}/{} {}:{:02}:{:02}", m, d, y, h, min, s)))
}

/// Date() / Today() - Returns current date as string
pub fn date_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if !args.is_empty() {
        return Err(RuntimeError::Custom("Date takes no arguments".to_string()));
    }
    let (y, m, d, _, _, _) = get_local_time();
    Ok(Value::String(format!("{}/{}/{}", m, d, y)))
}

/// Time() - Returns current time as string
pub fn time_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if !args.is_empty() {
        return Err(RuntimeError::Custom("Time takes no arguments".to_string()));
    }
    let (_, _, _, h, min, s) = get_local_time();
    Ok(Value::String(format!("{}:{:02}:{:02}", h, min, s)))
}

/// Year(date_string) - Extracts year from date string
pub fn year_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        let (y, _, _, _, _, _) = get_local_time();
        return Ok(Value::Integer(y));
    }
    let s = args[0].as_string();
    // Try to extract year from common formats: M/D/YYYY, YYYY-MM-DD
    if let Some(year) = parse_date_component(&s, DatePart::Year) {
        Ok(Value::Integer(year))
    } else {
        Err(RuntimeError::Custom(format!("Cannot extract Year from: {}", s)))
    }
}

/// Month(date_string) - Extracts month from date string
pub fn month_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        let (_, m, _, _, _, _) = get_local_time();
        return Ok(Value::Integer(m));
    }
    let s = args[0].as_string();
    if let Some(month) = parse_date_component(&s, DatePart::Month) {
        Ok(Value::Integer(month))
    } else {
        Err(RuntimeError::Custom(format!("Cannot extract Month from: {}", s)))
    }
}

/// Day(date_string) - Extracts day from date string
pub fn day_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        let (_, _, d, _, _, _) = get_local_time();
        return Ok(Value::Integer(d));
    }
    let s = args[0].as_string();
    if let Some(day) = parse_date_component(&s, DatePart::Day) {
        Ok(Value::Integer(day))
    } else {
        Err(RuntimeError::Custom(format!("Cannot extract Day from: {}", s)))
    }
}

/// Hour(time_string) - Extracts hour
pub fn hour_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        let (_, _, _, h, _, _) = get_local_time();
        return Ok(Value::Integer(h));
    }
    let s = args[0].as_string();
    // Extract from "H:MM:SS" or "M/D/YYYY H:MM:SS"
    let time_part = if s.contains(' ') { s.split(' ').last().unwrap_or("") } else { &s };
    let parts: Vec<&str> = time_part.split(':').collect();
    if let Some(h) = parts.first().and_then(|p| p.parse::<i32>().ok()) {
        Ok(Value::Integer(h))
    } else {
        Ok(Value::Integer(0))
    }
}

/// Minute(time_string) - Extracts minute
pub fn minute_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        let (_, _, _, _, m, _) = get_local_time();
        return Ok(Value::Integer(m));
    }
    let s = args[0].as_string();
    let time_part = if s.contains(' ') { s.split(' ').last().unwrap_or("") } else { &s };
    let parts: Vec<&str> = time_part.split(':').collect();
    if let Some(m) = parts.get(1).and_then(|p| p.parse::<i32>().ok()) {
        Ok(Value::Integer(m))
    } else {
        Ok(Value::Integer(0))
    }
}

/// Second(time_string) - Extracts second
pub fn second_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        let (_, _, _, _, _, s) = get_local_time();
        return Ok(Value::Integer(s));
    }
    let s = args[0].as_string();
    let time_part = if s.contains(' ') { s.split(' ').last().unwrap_or("") } else { &s };
    let parts: Vec<&str> = time_part.split(':').collect();
    if let Some(sec) = parts.get(2).and_then(|p| p.parse::<i32>().ok()) {
        Ok(Value::Integer(sec))
    } else {
        Ok(Value::Integer(0))
    }
}

/// Timer() - Returns seconds since midnight
pub fn timer_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if !args.is_empty() {
        return Err(RuntimeError::Custom("Timer takes no arguments".to_string()));
    }
    let secs = SystemTime::now()
        .duration_since(SystemTime::UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs();
    let time_of_day = secs % 86400;
    Ok(Value::Single(time_of_day as f32))
}

enum DatePart { Year, Month, Day }

fn parse_date_component(s: &str, part: DatePart) -> Option<i32> {
    // Try M/D/YYYY format
    let date_str = if s.contains(' ') { s.split(' ').next().unwrap_or(s) } else { s };

    if date_str.contains('/') {
        let parts: Vec<&str> = date_str.split('/').collect();
        if parts.len() == 3 {
            return match part {
                DatePart::Month => parts[0].parse().ok(),
                DatePart::Day => parts[1].parse().ok(),
                DatePart::Year => parts[2].parse().ok(),
            };
        }
    }
    // Try YYYY-MM-DD format
    if date_str.contains('-') {
        let parts: Vec<&str> = date_str.split('-').collect();
        if parts.len() == 3 {
            return match part {
                DatePart::Year => parts[0].parse().ok(),
                DatePart::Month => parts[1].parse().ok(),
                DatePart::Day => parts[2].parse().ok(),
            };
        }
    }
    None
}

/// DateAdd(interval, number, date) - Add time interval to date
pub fn dateadd_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    use chrono::Duration;
    
    if args.len() != 3 {
        return Err(RuntimeError::Custom("DateAdd requires 3 arguments".to_string()));
    }
    
    let interval = args[0].as_string();
    let number = args[1].as_integer()?;
    let date_str = args[2].as_string();
    
    // Parse the date
    let dt = parse_datetime(&date_str)?;
    
    // Add the interval
    let result = match interval.to_lowercase().as_str() {
        "yyyy" => dt + Duration::days(number as i64 * 365), // Approximate year
        "q" => dt + Duration::days(number as i64 * 91), // Approximate quarter
        "m" => dt + Duration::days(number as i64 * 30), // Approximate month
        "y" | "d" => dt + Duration::days(number as i64),
        "w" => dt + Duration::weeks(number as i64),
        "ww" => dt + Duration::weeks(number as i64),
        "h" => dt + Duration::hours(number as i64),
        "n" => dt + Duration::minutes(number as i64),
        "s" => dt + Duration::seconds(number as i64),
        _ => return Err(RuntimeError::Custom(format!("Invalid DateAdd interval: {}", interval))),
    };
    
    Ok(Value::String(result.format("%m/%d/%Y %H:%M:%S").to_string()))
}

/// DateDiff(interval, date1, date2) - Calculate difference between dates
pub fn datediff_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 3 {
        return Err(RuntimeError::Custom("DateDiff requires 3 arguments".to_string()));
    }
    
    let interval = args[0].as_string();
    let date1 = parse_datetime(&args[1].as_string())?;
    let date2 = parse_datetime(&args[2].as_string())?;
    
    let diff = date2.signed_duration_since(date1);
    
    let result = match interval.to_lowercase().as_str() {
        "yyyy" => diff.num_days() / 365,
        "q" => diff.num_days() / 91,
        "m" => diff.num_days() / 30,
        "y" | "d" => diff.num_days(),
        "w" => diff.num_weeks(),
        "ww" => diff.num_weeks(),
        "h" => diff.num_hours(),
        "n" => diff.num_minutes(),
        "s" => diff.num_seconds(),
        _ => return Err(RuntimeError::Custom(format!("Invalid DateDiff interval: {}", interval))),
    };
    
    Ok(Value::Integer(result as i32))
}

/// DatePart(interval, date) - Extract part of a date
pub fn datepart_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 2 {
        return Err(RuntimeError::Custom("DatePart requires 2 arguments".to_string()));
    }
    
    let interval = args[0].as_string();
    let date_str = args[1].as_string();
    let dt = parse_datetime(&date_str)?;
    
    let result = match interval.to_lowercase().as_str() {
        "yyyy" => dt.year(),
        "q" => ((dt.month() - 1) / 3) as i32 + 1,
        "m" => dt.month() as i32,
        "y" => dt.ordinal() as i32,
        "d" => dt.day() as i32,
        "w" => dt.weekday().num_days_from_sunday() as i32 + 1,
        "ww" => dt.iso_week().week() as i32,
        "h" => dt.hour() as i32,
        "n" => dt.minute() as i32,
        "s" => dt.second() as i32,
        _ => return Err(RuntimeError::Custom(format!("Invalid DatePart interval: {}", interval))),
    };
    
    Ok(Value::Integer(result))
}

/// DateSerial(year, month, day) - Create date from components
pub fn dateserial_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    use chrono::NaiveDate;
    
    if args.len() != 3 {
        return Err(RuntimeError::Custom("DateSerial requires 3 arguments".to_string()));
    }
    
    let year = args[0].as_integer()?;
    let month = args[1].as_integer()? as u32;
    let day = args[2].as_integer()? as u32;
    
    if let Some(date) = NaiveDate::from_ymd_opt(year, month, day) {
        Ok(Value::String(date.format("%m/%d/%Y").to_string()))
    } else {
        Err(RuntimeError::Custom(format!("Invalid date: {}/{}/{}", month, day, year)))
    }
}

/// TimeSerial(hour, minute, second) - Create time from components
pub fn timeserial_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 3 {
        return Err(RuntimeError::Custom("TimeSerial requires 3 arguments".to_string()));
    }
    
    let hour = args[0].as_integer()?;
    let minute = args[1].as_integer()?;
    let second = args[2].as_integer()?;
    
    Ok(Value::String(format!("{:02}:{:02}:{:02}", hour, minute, second)))
}

/// DateValue(date_string) - Parse date from string
pub fn datevalue_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("DateValue requires 1 argument".to_string()));
    }
    
    let date_str = args[0].as_string();
    let dt = parse_datetime(&date_str)?;
    Ok(Value::String(dt.format("%m/%d/%Y").to_string()))
}

/// TimeValue(time_string) - Parse time from string
pub fn timevalue_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("TimeValue requires 1 argument".to_string()));
    }
    
    let time_str = args[0].as_string();
    let dt = parse_datetime(&time_str)?;
    Ok(Value::String(dt.format("%H:%M:%S").to_string()))
}

/// MonthName(month, [abbreviate]) - Get month name
pub fn monthname_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() || args.len() > 2 {
        return Err(RuntimeError::Custom("MonthName requires 1 or 2 arguments".to_string()));
    }
    
    let month = args[0].as_integer()?;
    let abbrev = args.len() == 2 && args[1].as_bool().unwrap_or(false);
    
    let names = if abbrev {
        ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    } else {
        ["January", "February", "March", "April", "May", "June", 
         "July", "August", "September", "October", "November", "December"]
    };
    
    if month >= 1 && month <= 12 {
        Ok(Value::String(names[(month - 1) as usize].to_string()))
    } else {
        Err(RuntimeError::Custom(format!("Invalid month number: {}", month)))
    }
}

/// WeekdayName(weekday, [abbreviate], [firstdayofweek]) - Get weekday name
pub fn weekdayname_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() || args.len() > 3 {
        return Err(RuntimeError::Custom("WeekdayName requires 1 to 3 arguments".to_string()));
    }
    
    let weekday = args[0].as_integer()?;
    let abbrev = args.len() >= 2 && args[1].as_bool().unwrap_or(false);
    
    let names = if abbrev {
        ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
    } else {
        ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
    };
    
    if weekday >= 1 && weekday <= 7 {
        Ok(Value::String(names[(weekday - 1) as usize].to_string()))
    } else {
        Err(RuntimeError::Custom(format!("Invalid weekday number: {}", weekday)))
    }
}

/// Weekday(date, [firstdayofweek]) - Get weekday number
pub fn weekday_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() || args.len() > 2 {
        return Err(RuntimeError::Custom("Weekday requires 1 or 2 arguments".to_string()));
    }
    
    let date_str = args[0].as_string();
    let dt = parse_datetime(&date_str)?;
    
    // Sunday = 1, Monday = 2, etc.
    let day = dt.weekday().num_days_from_sunday() as i32 + 1;
    Ok(Value::Integer(day))
}

fn parse_datetime(s: &str) -> Result<chrono::NaiveDateTime, RuntimeError> {
    use chrono::{NaiveDate, NaiveTime, NaiveDateTime};
    
    let s = s.trim();
    
    // Try various formats
    let formats = [
        "%m/%d/%Y %H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%m/%d/%Y %H:%M",
        "%Y-%m-%d %H:%M",
        "%m/%d/%Y",
        "%Y-%m-%d",
        "%H:%M:%S",
        "%H:%M"
    ];
    
    for fmt in formats {
        if let Ok(dt) = NaiveDateTime::parse_from_str(s, fmt) {
            return Ok(dt);
        }
        if let Ok(d) = NaiveDate::parse_from_str(s, fmt) {
            return Ok(d.and_hms_opt(0, 0, 0).unwrap());
        }
        if let Ok(t) = NaiveTime::parse_from_str(s, fmt) {
            let base_date = NaiveDate::from_ymd_opt(1899, 12, 30).unwrap();
            return Ok(base_date.and_time(t));
        }
    }
    
    Err(RuntimeError::Custom(format!("Cannot parse date/time: {}", s)))
}
