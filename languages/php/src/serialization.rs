//! Decode PHP's wire serialization before the existing VM reconstruction step.
use serde_json::{Map, Value as Json, json};
use std::sync::Arc;
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{Framework, Value, heap};

pub fn register(fw: &mut Framework<'_>) {
    fw.register_host_fn(
        "php:serialization",
        "encodeScalarArray",
        Box::new(|_, args| {
            let Some(value) = args.first() else { return Value::Null; };
            let mut out = String::new();
            if encode_scalar_array(value, &mut out, 0).is_some() {
                Value::String(Arc::from(out))
            } else {
                // Objects with serialization hooks remain in the Wasm adapter.
                Value::Null
            }
        }),
    );
    fw.register_host_fn(
        "php:serialization",
        "assocSet",
        Box::new(|_, args| {
            let [Value::Object(array), key, value, ..] = args else {
                return Value::Null;
            };
            let key = match key {
                Value::String(text) => text
                    .parse::<i64>()
                    .ok()
                    .filter(|number| number.to_string() == text.as_ref())
                    .map(Value::I64)
                    .unwrap_or_else(|| key.clone()),
                _ => key.clone(),
            };
            let mut object = array.lock().unwrap();
            if let ObjectKind::Map(entries) = &mut object.kind {
                entries.insert(key, value.clone());
            }
            Value::Null
        }),
    );
    fw.register_host_fn(
        "php:serialization",
        "arrayToMap",
        Box::new(|_, args| {
            let Some(Value::Object(array)) = args.first() else {
                return Value::Null;
            };
            let items = {
                let object = array.lock().unwrap();
                let ObjectKind::Array(items) = &object.kind else {
                    return Value::Null;
                };
                items.clone()
            };
            let mut object = Object::new();
            object.kind = ObjectKind::Map(Default::default());
            if let ObjectKind::Map(entries) = &mut object.kind {
                for (index, value) in items.into_iter().enumerate() {
                    entries.insert(Value::I64(index as i64), value);
                }
            }
            Value::Object(heap::alloc(object))
        }),
    );
    fw.register_host_fn(
        "php:serialization",
        "toJson",
        Box::new(|_, args| {
            let text = match args.first() {
                Some(Value::String(s)) => s.as_ref(),
                _ => return Value::String(Arc::from("false")),
            };
            // Vybe's own serialize() currently emits JSON. Preserve that format
            // until the encoder is converted to PHP's wire format as well.
            if serde_json::from_str::<Json>(text).is_ok() {
                return Value::String(Arc::from(text));
            }
            let mut parser = Parser {
                bytes: text.as_bytes(),
                at: 0,
                depth: 0,
            };
            let value = parser.value();
            let encoded = match value {
                Some(value) if parser.at == parser.bytes.len() => value.to_string(),
                _ => "false".to_owned(),
            };
            Value::String(Arc::from(encoded))
        }),
    );
}

fn encode_scalar_array(value: &Value, out: &mut String, depth: usize) -> Option<()> {
    use std::fmt::Write;
    if depth >= 128 { return None; }
    match value {
        Value::Null => out.push_str("N;"),
        Value::Bool(value) => out.push_str(if *value { "b:1;" } else { "b:0;" }),
        Value::I32(value) => { write!(out, "i:{value};").ok()?; }
        Value::I64(value) => { write!(out, "i:{value};").ok()?; }
        Value::F64(value) if value.is_finite() && value.fract() == 0.0
            && *value >= i64::MIN as f64 && *value < i64::MAX as f64 => {
            write!(out, "i:{};", *value as i64).ok()?;
        }
        Value::F64(value) => {
            if value.is_nan() { out.push_str("d:NAN;"); }
            else if *value == f64::INFINITY { out.push_str("d:INF;"); }
            else if *value == f64::NEG_INFINITY { out.push_str("d:-INF;"); }
            else { write!(out, "d:{value};").ok()?; }
        }
        Value::String(value) => {
            write!(out, "s:{}:\"", value.len()).ok()?;
            out.push_str(value);
            out.push_str("\";");
        }
        Value::Object(object) => {
            // Release the lock before recursion, including cyclic input.
            let entries: Vec<(Value, Value)> = {
                let object = object.lock().unwrap();
                match &object.kind {
                    ObjectKind::Array(items) => items.iter().enumerate()
                        .map(|(index, item)| (Value::I64(index as i64), item.clone())).collect(),
                    ObjectKind::Map(items) => items.iter()
                        .map(|(key, item)| (key.clone(), item.clone())).collect(),
                    _ => return None,
                }
            };
            write!(out, "a:{}:{{", entries.len()).ok()?;
            for (key, item) in entries {
                if !matches!(key, Value::String(_) | Value::I32(_) | Value::I64(_) | Value::F64(_)) {
                    return None;
                }
                encode_scalar_array(&key, out, depth + 1)?;
                encode_scalar_array(&item, out, depth + 1)?;
            }
            out.push('}');
        }
        _ => return None,
    }
    Some(())
}

struct Parser<'a> {
    bytes: &'a [u8],
    at: usize,
    depth: usize,
}

impl Parser<'_> {
    fn take(&mut self, token: u8) -> Option<()> {
        if self.bytes.get(self.at).copied()? != token {
            return None;
        }
        self.at += 1;
        Some(())
    }

    fn until(&mut self, token: u8) -> Option<&[u8]> {
        let start = self.at;
        let end = self.bytes[start..].iter().position(|&b| b == token)? + start;
        self.at = end + 1;
        Some(&self.bytes[start..end])
    }

    fn usize(&mut self, end: u8) -> Option<usize> {
        std::str::from_utf8(self.until(end)?).ok()?.parse().ok()
    }

    fn string(&mut self) -> Option<String> {
        self.take(b's')?;
        self.take(b':')?;
        let len = self.usize(b':')?;
        self.take(b'"')?;
        let end = self.at.checked_add(len)?;
        let value = String::from_utf8(self.bytes.get(self.at..end)?.to_vec()).ok()?;
        self.at = end;
        self.take(b'"')?;
        self.take(b';')?;
        Some(value)
    }

    fn value(&mut self) -> Option<Json> {
        if self.depth >= 128 {
            return None;
        }
        self.depth += 1;
        let result = self.value_inner();
        self.depth -= 1;
        result
    }

    fn value_inner(&mut self) -> Option<Json> {
        match *self.bytes.get(self.at)? {
            b'N' => {
                self.at += 1;
                self.take(b';')?;
                Some(Json::Null)
            }
            b'b' => {
                self.at += 1;
                self.take(b':')?;
                let b = self.until(b';')?;
                Some(Json::Bool(b == b"1"))
            }
            b'i' => {
                self.at += 1;
                self.take(b':')?;
                let n = std::str::from_utf8(self.until(b';')?)
                    .ok()?
                    .parse::<i64>()
                    .ok()?;
                Some(json!(n))
            }
            b'd' => {
                self.at += 1;
                self.take(b':')?;
                let n = std::str::from_utf8(self.until(b';')?)
                    .ok()?
                    .parse::<f64>()
                    .ok()?;
                Some(json!(n))
            }
            b's' => Some(Json::String(self.string()?)),
            b'a' => {
                self.at += 1;
                self.take(b':')?;
                let len = self.usize(b':')?;
                if len > 1_000_000 {
                    return None;
                }
                self.take(b'{')?;
                let mut items = Vec::new();
                let mut assoc = Map::new();
                for _ in 0..len {
                    let key = self.value()?;
                    let value = self.value()?;
                    if key.as_i64() == Some(items.len() as i64) && assoc.is_empty() {
                        items.push(value);
                    } else {
                        let key = key
                            .as_str()
                            .map(str::to_owned)
                            .or_else(|| key.as_i64().map(|n| n.to_string()))?;
                        assoc.insert(key, value);
                    }
                }
                self.take(b'}')?;
                Some(json!({"vybe$php_ser_kind":"array", "items":items, "assoc":assoc}))
            }
            _ => None,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn parses_php_array_with_byte_counted_strings() {
        let mut p = Parser {
            bytes: b"a:2:{i:0;s:2:\"hi\";s:3:\"key\";b:1;}",
            at: 0,
            depth: 0,
        };
        let value = p.value().unwrap();
        assert_eq!(value["items"][0], "hi");
        assert_eq!(value["assoc"]["key"], true);
        assert_eq!(p.at, p.bytes.len());
    }
}
