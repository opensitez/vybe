use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value, HostContext};
use vybe_bytecode::value::Object;

pub fn register(vm: &mut VM) {
    // regex.test(pattern, string) → bool
    vm.register_host_fn("vybe:regex", "test", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let pattern = s(args, 0);
        let input = s(args, 1);
        match regex::Regex::new(&pattern) {
            Ok(re) => Value::Bool(re.is_match(&input)),
            Err(_) => Value::Bool(false),
        }
    }));

    // regex.match(pattern, string) → array of matches or null
    vm.register_host_fn("vybe:regex", "match", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let pattern = s(args, 0);
        let input = s(args, 1);
        match regex::Regex::new(&pattern) {
            Ok(re) => {
                let matches: Vec<Value> = re.find_iter(&input)
                    .map(|m| Value::String(Arc::from(m.as_str())))
                    .collect();
                if matches.is_empty() {
                    Value::Null
                } else {
                    Value::Object(Arc::new(Mutex::new(Object::new_array(matches))))
                }
            }
            Err(_) => Value::Null,
        }
    }));

    // regex.replace(pattern, string, replacement) → new string
    vm.register_host_fn("vybe:regex", "replace", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let pattern = s(args, 0);
        let input = s(args, 1);
        let replacement = s(args, 2);
        match regex::Regex::new(&pattern) {
            Ok(re) => Value::String(Arc::from(re.replace(&input, replacement.as_str()).as_ref())),
            Err(_) => Value::String(Arc::from(input.as_str())),
        }
    }));

    // regex.replaceAll(pattern, string, replacement) → new string
    vm.register_host_fn("vybe:regex", "replaceAll", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let pattern = s(args, 0);
        let input = s(args, 1);
        let replacement = s(args, 2);
        match regex::Regex::new(&pattern) {
            Ok(re) => Value::String(Arc::from(re.replace_all(&input, replacement.as_str()).as_ref())),
            Err(_) => Value::String(Arc::from(input.as_str())),
        }
    }));

    // regex.split(pattern, string) → array of parts
    vm.register_host_fn("vybe:regex", "split", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let pattern = s(args, 0);
        let input = s(args, 1);
        match regex::Regex::new(&pattern) {
            Ok(re) => {
                let parts: Vec<Value> = re.split(&input)
                    .map(|p| Value::String(Arc::from(p)))
                    .collect();
                Value::Object(Arc::new(Mutex::new(Object::new_array(parts))))
            }
            Err(_) => Value::Object(Arc::new(Mutex::new(Object::new_array(vec![
                Value::String(Arc::from(input.as_str()))
            ])))),
        }
    }));

    // regex.matchGroups(pattern, string) → object with named groups or array of capture groups
    vm.register_host_fn("vybe:regex", "matchGroups", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let pattern = s(args, 0);
        let input = s(args, 1);
        match regex::Regex::new(&pattern) {
            Ok(re) => {
                if let Some(caps) = re.captures(&input) {
                    let groups: Vec<Value> = caps.iter()
                        .map(|m| match m {
                            Some(m) => Value::String(Arc::from(m.as_str())),
                            None => Value::Null,
                        })
                        .collect();
                    Value::Object(Arc::new(Mutex::new(Object::new_array(groups))))
                } else {
                    Value::Null
                }
            }
            Err(_) => Value::Null,
        }
    }));
}

fn s(args: &[Value], idx: usize) -> String {
    args.get(idx).map(|v| format!("{}", v)).unwrap_or_default()
}

pub fn register_constructor(vm: &mut VM) {
    // New Regex(pattern[, flags]) → object { __pattern, __flags }
    vm.register_host_fn("vybe:types", "regexNew", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let pattern = s(args, 0);
        let flags = args.get(1).map(|v| v.as_f64() as i32).unwrap_or(0);
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Arc::from("Regex")));
        obj.properties.insert("__pattern".into(), Value::String(Arc::from(pattern.as_str())));
        obj.properties.insert("__flags".into(), Value::F64(flags as f64));
        Value::Object(Arc::new(Mutex::new(obj)))
    }));

    // Instance methods — extract __pattern + __flags from self object
    vm.register_host_fn("vybe:types", "regexIsMatch", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let (pat, case_insensitive) = get_regex_pattern(args.first());
        let input = s(args, 1);
        let full_pat = if case_insensitive { format!("(?i){}", pat) } else { pat };
        match regex::Regex::new(&full_pat) {
            Ok(re) => Value::Bool(re.is_match(&input)),
            Err(_) => Value::Bool(false),
        }
    }));

    vm.register_host_fn("vybe:types", "regexMatch", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let (pat, case_insensitive) = get_regex_pattern(args.first());
        let input = s(args, 1);
        let full_pat = if case_insensitive { format!("(?i){}", pat) } else { pat };
        match regex::Regex::new(&full_pat) {
            Ok(re) => match re.find(&input) {
                Some(m) => Value::String(Arc::from(m.as_str())),
                None => Value::Null,
            },
            Err(_) => Value::Null,
        }
    }));

    vm.register_host_fn("vybe:types", "regexMatches", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let (pat, case_insensitive) = get_regex_pattern(args.first());
        let input = s(args, 1);
        let full_pat = if case_insensitive { format!("(?i){}", pat) } else { pat };
        match regex::Regex::new(&full_pat) {
            Ok(re) => {
                let v: Vec<Value> = re.find_iter(&input)
                    .map(|m| Value::String(Arc::from(m.as_str())))
                    .collect();
                if v.is_empty() { Value::Null } else { Value::Object(Arc::new(Mutex::new(Object::new_array(v)))) }
            }
            Err(_) => Value::Null,
        }
    }));

    vm.register_host_fn("vybe:types", "regexReplace", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let (pat, case_insensitive) = get_regex_pattern(args.first());
        let input = s(args, 1);
        let replacement = s(args, 2);
        let full_pat = if case_insensitive { format!("(?i){}", pat) } else { pat };
        match regex::Regex::new(&full_pat) {
            Ok(re) => Value::String(Arc::from(re.replace_all(&input, replacement.as_str()).as_ref())),
            Err(_) => Value::String(Arc::from(input.as_str())),
        }
    }));
}

fn get_regex_pattern(val: Option<&Value>) -> (String, bool) {
    if let Some(Value::Object(obj)) = val {
        let o = obj.lock().unwrap();
        let pat = if let Some(Value::String(p)) = o.properties.get("__pattern") {
            p.to_string()
        } else {
            String::new()
        };
        let ci = if let Some(Value::F64(f)) = o.properties.get("__flags") {
            (*f as i32) & 1 != 0 // RegexOptions.IgnoreCase = 1
        } else {
            false
        };
        return (pat, ci);
    }
    (String::new(), false)
}
