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
        
        // Type checking functions
        "isnull" => {
            use crate::builtins::isnull_fn;
            isnull_fn(args)
        }
        
        // Conversion functions
        "ccur" => {
            use crate::builtins::ccur_fn;
            ccur_fn(args)
        }
        "cvar" => {
            use crate::builtins::cvar_fn;
            cvar_fn(args)
        }
        "csbyte" => {
            use crate::builtins::csbyte_fn;
            csbyte_fn(args)
        }
        
        // System objects
        "app" => {
            use crate::builtins::app_fn;
            app_fn(args)
        }
        "screen" => {
            use crate::builtins::screen_fn;
            screen_fn(args)
        }
        "clipboard" => {
            use crate::builtins::clipboard_fn;
            clipboard_fn(args)
        }
        "forms" => {
            use crate::builtins::forms_fn;
            forms_fn(args)
        }
        
        // System.Text
        "stringbuilder" => {
            use crate::builtins::stringbuilder_new_fn;
            stringbuilder_new_fn(args)
        }
        
        // Encoding
        "encoding.ascii.getbytes" => {
            use crate::builtins::encoding_ascii_getbytes_fn;
            encoding_ascii_getbytes_fn(args)
        }
        "encoding.ascii.getstring" => {
            use crate::builtins::encoding_ascii_getstring_fn;
            encoding_ascii_getstring_fn(args)
        }
        "encoding.utf8.getbytes" => {
            use crate::builtins::encoding_utf8_getbytes_fn;
            encoding_utf8_getbytes_fn(args)
        }
        "encoding.utf8.getstring" => {
            use crate::builtins::encoding_utf8_getstring_fn;
            encoding_utf8_getstring_fn(args)
        }
        "encoding.unicode.getbytes" => {
            use crate::builtins::encoding_unicode_getbytes_fn;
            encoding_unicode_getbytes_fn(args)
        }
        "encoding.unicode.getstring" => {
            use crate::builtins::encoding_unicode_getstring_fn;
            encoding_unicode_getstring_fn(args)
        }
        "encoding.default.getbytes" => {
            use crate::builtins::encoding_default_getbytes_fn;
            encoding_default_getbytes_fn(args)
        }
        "encoding.default.getstring" => {
            use crate::builtins::encoding_default_getstring_fn;
            encoding_default_getstring_fn(args)
        }
        "encoding.getencoding" => {
            use crate::builtins::encoding_getencoding_fn;
            encoding_getencoding_fn(args)
        }
        "encoding.convert" => {
            use crate::builtins::encoding_convert_fn;
            encoding_convert_fn(args)
        }
        
        // Regex
        "regex.ismatch" => {
            use crate::builtins::regex_ismatch_fn;
            regex_ismatch_fn(args)
        }
        "regex.match" => {
            use crate::builtins::regex_match_fn;
            regex_match_fn(args)
        }
        "regex.matches" => {
            use crate::builtins::regex_matches_fn;
            regex_matches_fn(args)
        }
        "regex.replace" => {
            use crate::builtins::regex_replace_fn;
            regex_replace_fn(args)
        }
        "regex.split" => {
            use crate::builtins::regex_split_fn;
            regex_split_fn(args)
        }
        
        // JSON
        "jsonserializer.serialize" | "json.serialize" => {
            use crate::builtins::json_serialize_fn;
            json_serialize_fn(args)
        }
        "jsonserializer.deserialize" | "json.deserialize" => {
            use crate::builtins::json_deserialize_fn;
            json_deserialize_fn(args)
        }
        
        // XML
        "xdocument.parse" | "xml.parse" => {
            use crate::builtins::xml_parse_fn;
            xml_parse_fn(args)
        }
        "xdocument.save" | "xml.save" => {
            use crate::builtins::xml_save_fn;
            xml_save_fn(args)
        }
        
        // File I/O
        "open" => {
            use crate::builtins::open_file_fn;
            open_file_fn(args)
        }
        "close" => {
            use crate::builtins::close_file_fn;
            close_file_fn(args)
        }
        "print" => {
            use crate::builtins::print_file_fn;
            print_file_fn(args)
        }
        "write" => {
            use crate::builtins::write_file_fn;
            write_file_fn(args)
        }
        "lineinput" => {
            use crate::builtins::line_input_fn;
            line_input_fn(args)
        }
        "seek" => {
            use crate::builtins::seek_file_fn;
            seek_file_fn(args)
        }
        "get" => {
            use crate::builtins::get_file_fn;
            get_file_fn(args)
        }
        "put" => {
            use crate::builtins::put_file_fn;
            put_file_fn(args)
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
