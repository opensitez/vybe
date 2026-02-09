use crate::value::{RuntimeError, Value};
use std::time::SystemTime;

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
