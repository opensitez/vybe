use crate::value::{Value, RuntimeError};
use chrono::{Local, Datelike, NaiveDate, Duration};

pub fn call_builtin(name: &str, args: &[Value]) -> Result<Value, RuntimeError> {
    match name.to_lowercase().as_str() {
        // Date Functions
        "now" => Ok(Value::Date(get_ole_now())),
        "date" => Ok(Value::Date(get_ole_date())),
        "time" => Ok(Value::Date(get_ole_time())),
        "year" => {
            let d = args.get(0).ok_or(RuntimeError::Custom("Argument required".to_string()))?.as_double()?;
            Ok(Value::Integer(get_date_part(d, DatePart::Year)))
        }
        "month" => {
            let d = args.get(0).ok_or(RuntimeError::Custom("Argument required".to_string()))?.as_double()?;
            Ok(Value::Integer(get_date_part(d, DatePart::Month)))
        }
        "day" => {
            let d = args.get(0).ok_or(RuntimeError::Custom("Argument required".to_string()))?.as_double()?;
            Ok(Value::Integer(get_date_part(d, DatePart::Day)))
        }
        
        // String Functions
        "len" => {
            let s = args.get(0).ok_or(RuntimeError::Custom("Argument required".to_string()))?.as_string();
            Ok(Value::Integer(s.len() as i32))
        }
        "mid" => {
            let s = args.get(0).ok_or(RuntimeError::Custom("Argument required".to_string()))?.as_string();
            let start = args.get(1).ok_or(RuntimeError::Custom("Start required".to_string()))?.as_integer()? as usize;
            let length = if let Some(len_arg) = args.get(2) {
                Some(len_arg.as_integer()? as usize)
            } else {
                None
            };
            
            if start == 0 { return Err(RuntimeError::Custom("Mid start must be > 0".to_string())); }
            let chars: Vec<char> = s.chars().collect();
            if start > chars.len() {
                Ok(Value::String(String::new()))
            } else {
                let real_start = start - 1;
                let real_len = length.unwrap_or(chars.len() - real_start);
                let end = std::cmp::min(real_start + real_len, chars.len());
                Ok(Value::String(chars[real_start..end].iter().collect()))
            }
        }
        
        // Math Functions
        "abs" => {
            let val = args.get(0).ok_or(RuntimeError::Custom("Argument required".to_string()))?;
            match val {
                Value::Integer(i) => Ok(Value::Integer(i.abs())),
                Value::Double(d) => Ok(Value::Double(d.abs())),
                _ => Ok(Value::Double(val.as_double()?.abs()))
            }
        }
        "round" => {
            let d = args.get(0).ok_or(RuntimeError::Custom("Argument required".to_string()))?.as_double()?;
            Ok(Value::Double(d.round()))
        }
        
        _ => Err(RuntimeError::UndefinedFunction(name.to_string())),
    }
}

// Helpers

enum DatePart {
    Year,
    Month,
    Day,
}

fn get_ole_now() -> f64 {
    let now = Local::now();
    let base_date = NaiveDate::from_ymd_opt(1899, 12, 30).unwrap().and_hms_opt(0, 0, 0).unwrap();
    let duration = now.naive_local().signed_duration_since(base_date);
    
    let days = duration.num_days() as f64;
    let seconds_in_day = (duration.num_seconds() % 86400) as f64;
    days + (seconds_in_day / 86400.0)
}

fn get_ole_date() -> f64 {
    let now = get_ole_now();
    now.trunc()
}

fn get_ole_time() -> f64 {
    let now = get_ole_now();
    now.fract()
}

fn get_date_part(ole_date: f64, part: DatePart) -> i32 {
    let base_date = NaiveDate::from_ymd_opt(1899, 12, 30).unwrap().and_hms_opt(0, 0, 0).unwrap();
    let days = ole_date.trunc() as i64;
    let fraction = ole_date.fract();
    let seconds = (fraction * 86400.0).round() as i64;
    
    if let Some(date) = base_date.checked_add_signed(Duration::days(days)) {
         if let Some(final_date) = date.checked_add_signed(Duration::seconds(seconds)) {
             return match part {
                 DatePart::Year => final_date.year(),
                 DatePart::Month => final_date.month() as i32,
                 DatePart::Day => final_date.day() as i32,
             };
         }
    }
    0
}
