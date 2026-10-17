use std::borrow::Cow;
use vybe_runtime::Value;
fn text(v:Option<&Value>)->Cow<'_,str>{
 match v {
  None=>Cow::Borrowed(""),Some(Value::String(s))=>Cow::Borrowed(s),
  Some(Value::I32(n))=>Cow::Owned(n.to_string()),Some(Value::I64(n))=>Cow::Owned(n.to_string()),Some(Value::F32(n))=>Cow::Owned(n.to_string()),Some(Value::F64(n))=>Cow::Owned(n.to_string()),
  Some(Value::Null)|Some(Value::TypedNull(_))=>Cow::Borrowed("null"),Some(Value::Undefined)=>Cow::Borrowed("undefined"),Some(Value::Bool(true))=>Cow::Borrowed("true"),Some(Value::Bool(false))=>Cow::Borrowed("false"),
  Some(Value::BigInt(n))=>Cow::Owned(n.to_string()),Some(other)=>Cow::Owned(vybe_platform_ecma::keys::value_display_string(other)),
 }
}
fn to_integer_or_infinity(value: Option<&Value>, default_for_undefined: f64) -> f64 {
    let number = match value {
        None | Some(Value::Undefined) => default_for_undefined,
        Some(Value::Null) => 0.0,
        Some(Value::Bool(v)) => {
            if *v {
                1.0
            } else {
                0.0
            }
        }
        Some(Value::I32(n)) => *n as f64,
        Some(Value::I64(n)) => *n as f64,
        Some(Value::F32(n)) => *n as f64,
        Some(Value::F64(n)) => *n,
        Some(Value::String(text)) => {
            let trimmed = text.trim();
            if trimmed.is_empty() {
                0.0
            } else if trimmed == "Infinity" || trimmed == "+Infinity" {
                f64::INFINITY
            } else if trimmed == "-Infinity" {
                f64::NEG_INFINITY
            } else {
                trimmed.parse::<f64>().unwrap_or(f64::NAN)
            }
        }
        Some(other) => format!("{other}").parse::<f64>().unwrap_or(f64::NAN),
    };
    if number.is_nan() || number == 0.0 {
        0.0
    } else if number.is_infinite() {
        number
    } else {
        number.trunc()
    }
}

pub fn invoke(args:&[Value])->Value {
 let source=text(args.first());let needle=text(args.get(1));let pos=to_integer_or_infinity(args.get(2),f64::INFINITY);
 if source.is_ascii()&&needle.is_ascii(){let end=if pos.is_infinite()&&pos.is_sign_positive(){source.len()}else if pos<=0.0{0}else if pos>=source.len()as f64{source.len()}else{pos as usize};if needle.is_empty(){return Value::F64(end as f64);}if needle.len()>source.len(){return Value::F64(-1.0);}let bound=end.min(source.len()-needle.len());return Value::F64(source[..bound+needle.len()].rfind(needle.as_ref()).map(|i|i as f64).unwrap_or(-1.0));}
 let hay:Vec<_>=source.encode_utf16().collect();let ndl:Vec<_>=needle.encode_utf16().collect();let end=if pos.is_infinite()&&pos.is_sign_positive(){hay.len()}else if pos<=0.0{0}else if pos>=hay.len()as f64{hay.len()}else{pos as usize};
 if ndl.is_empty(){return Value::F64(end as f64);}if ndl.len()>hay.len(){return Value::F64(-1.0);}let start=end.min(hay.len()-ndl.len());for i in(0..=start).rev(){if hay[i..i+ndl.len()]==ndl[..]{return Value::F64(i as f64);}}Value::F64(-1.0)
}
