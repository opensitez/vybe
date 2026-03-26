use std::cell::RefCell;
use std::rc::Rc;
use vybe_bytecode::{VM, Value};
use vybe_bytecode::value::Object;

pub fn register(vm: &mut VM) {
    // regex.test(pattern, string) → bool
    vm.register_host_fn("vybe:regex", "test", Box::new(|args: &[Value]| {
        let pattern = s(args, 0);
        let input = s(args, 1);
        match regex::Regex::new(&pattern) {
            Ok(re) => Value::Bool(re.is_match(&input)),
            Err(_) => Value::Bool(false),
        }
    }));

    // regex.match(pattern, string) → array of matches or null
    vm.register_host_fn("vybe:regex", "match", Box::new(|args: &[Value]| {
        let pattern = s(args, 0);
        let input = s(args, 1);
        match regex::Regex::new(&pattern) {
            Ok(re) => {
                let matches: Vec<Value> = re.find_iter(&input)
                    .map(|m| Value::String(Rc::from(m.as_str())))
                    .collect();
                if matches.is_empty() {
                    Value::Null
                } else {
                    Value::Object(Rc::new(RefCell::new(Object::new_array(matches))))
                }
            }
            Err(_) => Value::Null,
        }
    }));

    // regex.replace(pattern, string, replacement) → new string
    vm.register_host_fn("vybe:regex", "replace", Box::new(|args: &[Value]| {
        let pattern = s(args, 0);
        let input = s(args, 1);
        let replacement = s(args, 2);
        match regex::Regex::new(&pattern) {
            Ok(re) => Value::String(Rc::from(re.replace(&input, replacement.as_str()).as_ref())),
            Err(_) => Value::String(Rc::from(input.as_str())),
        }
    }));

    // regex.replaceAll(pattern, string, replacement) → new string
    vm.register_host_fn("vybe:regex", "replaceAll", Box::new(|args: &[Value]| {
        let pattern = s(args, 0);
        let input = s(args, 1);
        let replacement = s(args, 2);
        match regex::Regex::new(&pattern) {
            Ok(re) => Value::String(Rc::from(re.replace_all(&input, replacement.as_str()).as_ref())),
            Err(_) => Value::String(Rc::from(input.as_str())),
        }
    }));

    // regex.split(pattern, string) → array of parts
    vm.register_host_fn("vybe:regex", "split", Box::new(|args: &[Value]| {
        let pattern = s(args, 0);
        let input = s(args, 1);
        match regex::Regex::new(&pattern) {
            Ok(re) => {
                let parts: Vec<Value> = re.split(&input)
                    .map(|p| Value::String(Rc::from(p)))
                    .collect();
                Value::Object(Rc::new(RefCell::new(Object::new_array(parts))))
            }
            Err(_) => Value::Object(Rc::new(RefCell::new(Object::new_array(vec![
                Value::String(Rc::from(input.as_str()))
            ])))),
        }
    }));

    // regex.matchGroups(pattern, string) → object with named groups or array of capture groups
    vm.register_host_fn("vybe:regex", "matchGroups", Box::new(|args: &[Value]| {
        let pattern = s(args, 0);
        let input = s(args, 1);
        match regex::Regex::new(&pattern) {
            Ok(re) => {
                if let Some(caps) = re.captures(&input) {
                    let groups: Vec<Value> = caps.iter()
                        .map(|m| match m {
                            Some(m) => Value::String(Rc::from(m.as_str())),
                            None => Value::Null,
                        })
                        .collect();
                    Value::Object(Rc::new(RefCell::new(Object::new_array(groups))))
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
