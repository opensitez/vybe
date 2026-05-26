//! WHATWG URL Living Standard — URL + URLSearchParams.

use std::sync::{Arc, Mutex, OnceLock};
use url::{Origin, Url};
use vybe_bytecode::{HostContext, VM, Value};
use vybe_bytecode::value::{Object, ObjectKind};

const MODULE: &str = "web:url";

#[derive(Clone, Copy)]
struct HostFns {
    url_to_string: usize,
    url_to_json: usize,
    set_href: usize,
    set_protocol: usize,
    set_username: usize,
    set_password: usize,
    set_host: usize,
    set_hostname: usize,
    set_port: usize,
    set_pathname: usize,
    set_search: usize,
    set_hash: usize,
    params_get: usize,
    params_get_all: usize,
    params_has: usize,
    params_to_string: usize,
    params_append: usize,
    params_set: usize,
    params_delete: usize,
    params_sort: usize,
    params_keys: usize,
    params_values: usize,
    params_entries: usize,
    params_iterator: usize,
    params_for_each: usize,
}

static HOST_FNS: OnceLock<HostFns> = OnceLock::new();

fn str_value(value: impl Into<String>) -> Value {
    let value = value.into();
    Value::String(Arc::<str>::from(value))
}

fn host_fn_ref_by_idx(module: &str, name: &str, idx: usize) -> Value {
    let mut obj = Object::new();
    obj.properties.insert("__host_module".into(), str_value(module));
    obj.properties.insert("__host_name".into(), str_value(name));
    obj.properties.insert("__host_idx".into(), Value::F64(idx as f64));
    obj.properties.insert("name".into(), str_value(name));
    obj.kind = ObjectKind::HostFunction(idx);
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn bound_host_fn_ref_by_idx(module: &str, name: &str, idx: usize, bound_args: Vec<Value>) -> Value {
    let mut obj = Object::new();
    obj.properties.insert("__host_module".into(), str_value(module));
    obj.properties.insert("__host_name".into(), str_value(name));
    obj.properties.insert("__host_idx".into(), Value::F64(idx as f64));
    obj.properties.insert("name".into(), str_value(name));
    obj.properties.insert(
        "__bound_args".into(),
        Value::Object(Arc::new(Mutex::new(Object::new_array(bound_args)))),
    );
    obj.kind = ObjectKind::HostFunction(idx);
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn receiver_fn(name: &str, idx: usize) -> Value {
    crate::namespaces::receiver_host_fn_ref(MODULE, name, idx)
}

fn host_fns() -> &'static HostFns {
    HOST_FNS.get().expect("web:url host fns registered")
}

fn parse_url(input: &str, base: Option<&str>) -> Option<Url> {
    let input = input.trim();
    if let Ok(url) = Url::parse(input) {
        return Some(url);
    }
    let base = base?.trim();
    let base_url = Url::parse(base).ok()?;
    base_url.join(input).ok()
}

fn current_url(obj: &Arc<Mutex<Object>>) -> Option<Url> {
    let href = {
        let lock = obj.lock().unwrap();
        lock.properties.get("href").map(|v| format!("{}", v))
    }?;
    Url::parse(&href).ok()
}

fn origin_string(url: &Url) -> String {
    match url.origin() {
        Origin::Opaque(_) => "null".into(),
        origin => origin.ascii_serialization(),
    }
}

fn parse_query_pairs(query: &str) -> Vec<(String, String)> {
    form_urlencoded::parse(query.trim_start_matches('?').as_bytes())
        .into_owned()
        .collect()
}

fn serialize_pairs(pairs: &[(String, String)]) -> String {
    let mut serializer = form_urlencoded::Serializer::new(String::new());
    for (key, value) in pairs {
        serializer.append_pair(key, value);
    }
    serializer.finish()
}

fn pair_value(key: &str, value: &str) -> Value {
    Value::Object(Arc::new(Mutex::new(Object::new_array(vec![
        str_value(key),
        str_value(value),
    ]))))
}

fn pairs_from_params(obj: &Arc<Mutex<Object>>) -> Vec<(String, String)> {
    let pairs_value = {
        let lock = obj.lock().unwrap();
        lock.properties.get("__pairs").cloned()
    };
    let Some(Value::Object(pairs_obj)) = pairs_value else {
        return Vec::new();
    };
    let pairs_lock = pairs_obj.lock().unwrap();
    let ObjectKind::Array(pairs) = &pairs_lock.kind else {
        return Vec::new();
    };
    pairs
        .iter()
        .filter_map(|pair| match pair {
            Value::Object(pair_obj) => {
                let pair_lock = pair_obj.lock().unwrap();
                let ObjectKind::Array(values) = &pair_lock.kind else {
                    return None;
                };
                Some((
                    values.first().map(|v| format!("{}", v)).unwrap_or_default(),
                    values.get(1).map(|v| format!("{}", v)).unwrap_or_default(),
                ))
            }
            _ => None,
        })
        .collect()
}

fn sync_pairs_object(obj: &Arc<Mutex<Object>>, pairs: &[(String, String)]) {
    let entries: Vec<Value> = pairs.iter().map(|(key, value)| pair_value(key, value)).collect();
    let mut lock = obj.lock().unwrap();
    match lock.properties.get("__pairs").cloned() {
        Some(Value::Object(pairs_obj)) => {
            pairs_obj.lock().unwrap().kind = ObjectKind::Array(entries);
        }
        _ => {
            lock.properties.insert(
                "__pairs".into(),
                Value::Object(Arc::new(Mutex::new(Object::new_array(entries)))),
            );
        }
    }
    lock.properties.insert("size".into(), Value::F64(pairs.len() as f64));
}

fn split_host_port(input: &str) -> (String, Option<String>) {
    if input.starts_with('[') {
        if let Some(end) = input.rfind(']') {
            let host = input[..=end].to_string();
            let port = input[end + 1..]
                .strip_prefix(':')
                .filter(|port| !port.is_empty())
                .map(ToOwned::to_owned);
            return (host, port);
        }
    }
    if let Some((host, port)) = input.rsplit_once(':') {
        if !port.is_empty() && port.chars().all(|ch| ch.is_ascii_digit()) {
            return (host.to_string(), Some(port.to_string()));
        }
    }
    (input.to_string(), None)
}

fn coerce_init_pairs(value: &Value) -> Vec<(String, String)> {
    match value {
        Value::Undefined | Value::Null => Vec::new(),
        Value::Object(obj) => {
            let lock = obj.lock().unwrap();
            if let Some(Value::String(kind)) = lock.properties.get("__type") {
                if &**kind == "URLSearchParams" {
                    drop(lock);
                    return pairs_from_params(obj);
                }
            }
            if let ObjectKind::Array(items) = &lock.kind {
                let items = items.clone();
                drop(lock);
                return items
                    .iter()
                    .filter_map(|item| match item {
                        Value::Object(pair_obj) => {
                            let pair_lock = pair_obj.lock().unwrap();
                            let ObjectKind::Array(values) = &pair_lock.kind else {
                                return None;
                            };
                            Some((
                                values.first().map(|v| format!("{}", v)).unwrap_or_default(),
                                values.get(1).map(|v| format!("{}", v)).unwrap_or_default(),
                            ))
                        }
                        _ => None,
                    })
                    .collect();
            }
            let props: Vec<(String, Value)> = lock
                .properties
                .iter()
                .filter(|(key, _)| !key.starts_with("__") && *key != "size")
                .map(|(key, value)| (key.clone(), value.clone()))
                .collect();
            drop(lock);
            props.into_iter().map(|(key, value)| (key, format!("{}", value))).collect()
        }
        _ => parse_query_pairs(&format!("{}", value)),
    }
}

fn make_search_params_object(pairs: Vec<(String, String)>, owner: Option<Arc<Mutex<Object>>>) -> Value {
    let fns = host_fns();
    let mut obj = Object::new();
    obj.properties.insert("__type".into(), str_value("URLSearchParams"));
    obj.properties.insert("get".into(), receiver_fn("searchParamsGet", fns.params_get));
    obj.properties.insert("getAll".into(), receiver_fn("searchParamsGetAll", fns.params_get_all));
    obj.properties.insert("has".into(), receiver_fn("searchParamsHas", fns.params_has));
    obj.properties.insert(
        "toString".into(),
        bound_host_fn_ref_by_idx(MODULE, "searchParamsToString", fns.params_to_string, Vec::new()),
    );
    obj.properties.insert("append".into(), receiver_fn("searchParamsAppend", fns.params_append));
    obj.properties.insert("set".into(), receiver_fn("searchParamsSet", fns.params_set));
    obj.properties.insert("delete".into(), receiver_fn("searchParamsDelete", fns.params_delete));
    obj.properties.insert("sort".into(), receiver_fn("searchParamsSort", fns.params_sort));
    obj.properties.insert("keys".into(), receiver_fn("searchParamsKeys", fns.params_keys));
    obj.properties.insert("values".into(), receiver_fn("searchParamsValues", fns.params_values));
    obj.properties.insert("entries".into(), receiver_fn("searchParamsEntries", fns.params_entries));
    obj.properties.insert(
        "iterator".into(),
        bound_host_fn_ref_by_idx(MODULE, "searchParamsIterator", fns.params_iterator, Vec::new()),
    );
    obj.properties.insert("forEach".into(), receiver_fn("searchParamsForEach", fns.params_for_each));
    if let Some(owner) = owner {
        obj.properties.insert("__url_owner".into(), Value::Object(owner));
    }
    let value = Value::Object(Arc::new(Mutex::new(obj)));
    if let Value::Object(params_obj) = &value {
        let mut lock = params_obj.lock().unwrap();
        if matches!(lock.properties.get("toString"), Some(Value::Object(method)) if matches!(method.lock().unwrap().kind, ObjectKind::HostFunction(_))) {
            lock.properties.insert(
                "toString".into(),
                bound_host_fn_ref_by_idx(MODULE, "searchParamsToString", fns.params_to_string, vec![Value::Object(params_obj.clone())]),
            );
            lock.properties.insert(
                "iterator".into(),
                bound_host_fn_ref_by_idx(MODULE, "searchParamsIterator", fns.params_iterator, vec![Value::Object(params_obj.clone())]),
            );
        }
        drop(lock);
        sync_pairs_object(params_obj, &pairs);
    }
    value
}

fn sync_search_params_owner(params_obj: &Arc<Mutex<Object>>, owner: &Arc<Mutex<Object>>) {
    params_obj
        .lock()
        .unwrap()
        .properties
        .insert("__url_owner".into(), Value::Object(owner.clone()));
}

fn sync_url_object(obj: &Arc<Mutex<Object>>, url: &Url) {
    let hostname = url.host_str().unwrap_or_default().to_string();
    let port = url.port().map(|value| value.to_string()).unwrap_or_default();
    let host = if hostname.is_empty() {
        String::new()
    } else if port.is_empty() {
        hostname.clone()
    } else {
        format!("{}:{}", hostname, port)
    };
    let search = url.query().map(|query| format!("?{}", query)).unwrap_or_default();
    let hash = url.fragment().map(|fragment| format!("#{}", fragment)).unwrap_or_default();
    let pathname = if url.path().is_empty() { "/".into() } else { url.path().to_string() };
    let search_pairs = parse_query_pairs(&search);
    let existing_search_params = {
        let lock = obj.lock().unwrap();
        match lock.properties.get("searchParams") {
            Some(Value::Object(existing)) => Some(existing.clone()),
            _ => None,
        }
    };

    {
        let mut lock = obj.lock().unwrap();
        lock.properties.insert("__type".into(), str_value("URL"));
        lock.properties.insert("href".into(), str_value(url.to_string()));
        lock.properties.insert("origin".into(), str_value(origin_string(url)));
        lock.properties.insert("protocol".into(), str_value(format!("{}:", url.scheme())));
        lock.properties.insert("username".into(), str_value(url.username()));
        lock.properties.insert("password".into(), str_value(url.password().unwrap_or_default()));
        lock.properties.insert("host".into(), str_value(host));
        lock.properties.insert("hostname".into(), str_value(hostname));
        lock.properties.insert("port".into(), str_value(port));
        lock.properties.insert("pathname".into(), str_value(pathname));
        lock.properties.insert("search".into(), str_value(search));
        lock.properties.insert("hash".into(), str_value(hash));
        let fns = host_fns();
        lock.properties.insert(
            "toString".into(),
            bound_host_fn_ref_by_idx(MODULE, "urlToString", fns.url_to_string, vec![Value::Object(obj.clone())]),
        );
        lock.properties.insert(
            "toJSON".into(),
            bound_host_fn_ref_by_idx(MODULE, "urlToJSON", fns.url_to_json, vec![Value::Object(obj.clone())]),
        );
        lock.properties.insert("__set_href".into(), host_fn_ref_by_idx(MODULE, "urlSetHref", fns.set_href));
        lock.properties.insert("__set_protocol".into(), host_fn_ref_by_idx(MODULE, "urlSetProtocol", fns.set_protocol));
        lock.properties.insert("__set_username".into(), host_fn_ref_by_idx(MODULE, "urlSetUsername", fns.set_username));
        lock.properties.insert("__set_password".into(), host_fn_ref_by_idx(MODULE, "urlSetPassword", fns.set_password));
        lock.properties.insert("__set_host".into(), host_fn_ref_by_idx(MODULE, "urlSetHost", fns.set_host));
        lock.properties.insert("__set_hostname".into(), host_fn_ref_by_idx(MODULE, "urlSetHostname", fns.set_hostname));
        lock.properties.insert("__set_port".into(), host_fn_ref_by_idx(MODULE, "urlSetPort", fns.set_port));
        lock.properties.insert("__set_pathname".into(), host_fn_ref_by_idx(MODULE, "urlSetPathname", fns.set_pathname));
        lock.properties.insert("__set_search".into(), host_fn_ref_by_idx(MODULE, "urlSetSearch", fns.set_search));
        lock.properties.insert("__set_hash".into(), host_fn_ref_by_idx(MODULE, "urlSetHash", fns.set_hash));
    }

    let search_params = match existing_search_params {
        Some(existing) => {
            sync_pairs_object(&existing, &search_pairs);
            sync_search_params_owner(&existing, obj);
            Value::Object(existing)
        }
        None => make_search_params_object(search_pairs, Some(obj.clone())),
    };

    obj.lock().unwrap().properties.insert("searchParams".into(), search_params);
}

fn make_url_object(url: &Url) -> Value {
    let value = Value::Object(Arc::new(Mutex::new(Object::new())));
    if let Value::Object(obj) = &value {
        sync_url_object(obj, url);
    }
    value
}

fn update_owner_from_params(params_obj: &Arc<Mutex<Object>>) {
    let owner = {
        let lock = params_obj.lock().unwrap();
        match lock.properties.get("__url_owner") {
            Some(Value::Object(owner)) => Some(owner.clone()),
            _ => None,
        }
    };
    let Some(owner) = owner else {
        return;
    };
    let Some(mut url) = current_url(&owner) else {
        return;
    };
    let query = serialize_pairs(&pairs_from_params(params_obj));
    if query.is_empty() {
        url.set_query(None);
    } else {
        url.set_query(Some(&query));
    }
    sync_url_object(&owner, &url);
}

fn mutate_url_property<F>(args: &[Value], mutator: F) -> Value
where
    F: FnOnce(&mut Url, &str),
{
    let Some(Value::Object(obj)) = args.first() else {
        return Value::Undefined;
    };
    let Some(mut url) = current_url(obj) else {
        return Value::Undefined;
    };
    let value = args.get(1).map(|value| format!("{}", value)).unwrap_or_default();
    mutator(&mut url, &value);
    sync_url_object(obj, &url);
    Value::Undefined
}

fn params_array(values: Vec<Value>) -> Value {
    Value::Object(Arc::new(Mutex::new(Object::new_array(values))))
}

fn make_bound_array_iterator(values: Vec<Value>) -> Value {
    let iterator = crate::ecma::array::make_array_iterator(values);
    if let Value::Object(iterator_obj) = &iterator {
        let next_idx = {
            let lock = iterator_obj.lock().unwrap();
            match lock.properties.get("next") {
                Some(Value::Object(next)) => match next.lock().unwrap().properties.get("__host_idx") {
                    Some(Value::F64(idx)) => Some(*idx as usize),
                    _ => None,
                },
                _ => None,
            }
        };
        if let Some(next_idx) = next_idx {
            iterator_obj.lock().unwrap().properties.insert(
                "next".into(),
                bound_host_fn_ref_by_idx("ecma:array", "iterNext", next_idx, vec![Value::Object(iterator_obj.clone())]),
            );
        }
    }
    iterator
}

pub fn register(vm: &mut VM) {
    vm.register_host_fn(MODULE, "new", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let input = args.first().map(|value| format!("{}", value)).unwrap_or_default();
        let base = args.get(1).map(|value| format!("{}", value));
        parse_url(&input, base.as_deref())
            .map(|url| make_url_object(&url))
            .unwrap_or(Value::Null)
    }));

    vm.register_host_fn(MODULE, "parse", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let input = args.first().map(|value| format!("{}", value)).unwrap_or_default();
        let base = args.get(1).map(|value| format!("{}", value));
        parse_url(&input, base.as_deref())
            .map(|url| make_url_object(&url))
            .unwrap_or(Value::Null)
    }));

    vm.register_host_fn(MODULE, "canParse", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let input = args.first().map(|value| format!("{}", value)).unwrap_or_default();
        let base = args.get(1).map(|value| format!("{}", value));
        Value::Bool(parse_url(&input, base.as_deref()).is_some())
    }));

    vm.register_host_fn(MODULE, "urlToString", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        match args.first() {
            Some(Value::Object(obj)) => obj.lock().unwrap().properties.get("href").cloned().unwrap_or_else(|| str_value("")),
            _ => str_value(""),
        }
    }));

    vm.register_host_fn(MODULE, "urlToJSON", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        match args.first() {
            Some(Value::Object(obj)) => obj.lock().unwrap().properties.get("href").cloned().unwrap_or_else(|| str_value("")),
            _ => str_value(""),
        }
    }));

    vm.register_host_fn(MODULE, "urlSetHref", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(Value::Object(obj)) = args.first() else {
            return Value::Undefined;
        };
        let input = args.get(1).map(|value| format!("{}", value)).unwrap_or_default();
        if let Some(url) = parse_url(&input, None) {
            sync_url_object(obj, &url);
        }
        Value::Undefined
    }));

    vm.register_host_fn(MODULE, "urlSetProtocol", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        mutate_url_property(args, |url, value| {
            let scheme = value.trim().trim_end_matches(':').to_ascii_lowercase();
            if !scheme.is_empty() {
                let _ = url.set_scheme(&scheme);
            }
        })
    }));

    vm.register_host_fn(MODULE, "urlSetUsername", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        mutate_url_property(args, |url, value| {
            let _ = url.set_username(value);
        })
    }));

    vm.register_host_fn(MODULE, "urlSetPassword", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        mutate_url_property(args, |url, value| {
            let _ = if value.is_empty() {
                url.set_password(None)
            } else {
                url.set_password(Some(value))
            };
        })
    }));

    vm.register_host_fn(MODULE, "urlSetHost", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        mutate_url_property(args, |url, value| {
            let (host, port) = split_host_port(value);
            if !host.is_empty() {
                let _ = url.set_host(Some(&host));
            }
            match port.and_then(|port| port.parse::<u16>().ok()) {
                Some(port) => {
                    let _ = url.set_port(Some(port));
                }
                None => {
                    let _ = url.set_port(None);
                }
            }
        })
    }));

    vm.register_host_fn(MODULE, "urlSetHostname", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        mutate_url_property(args, |url, value| {
            if !value.is_empty() {
                let _ = url.set_host(Some(value));
            }
        })
    }));

    vm.register_host_fn(MODULE, "urlSetPort", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        mutate_url_property(args, |url, value| {
            let trimmed = value.trim();
            if trimmed.is_empty() {
                let _ = url.set_port(None);
            } else if let Ok(port) = trimmed.parse::<u16>() {
                let _ = url.set_port(Some(port));
            }
        })
    }));

    vm.register_host_fn(MODULE, "urlSetPathname", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        mutate_url_property(args, |url, value| {
            let pathname = if value.is_empty() {
                "/".to_string()
            } else if value.starts_with('/') {
                value.to_string()
            } else {
                format!("/{}", value)
            };
            url.set_path(&pathname);
        })
    }));

    vm.register_host_fn(MODULE, "urlSetSearch", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        mutate_url_property(args, |url, value| {
            let query = value.trim().trim_start_matches('?');
            if query.is_empty() {
                url.set_query(None);
            } else {
                url.set_query(Some(query));
            }
        })
    }));

    vm.register_host_fn(MODULE, "urlSetHash", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        mutate_url_property(args, |url, value| {
            let fragment = value.trim().trim_start_matches('#');
            if fragment.is_empty() {
                url.set_fragment(None);
            } else {
                url.set_fragment(Some(fragment));
            }
        })
    }));

    vm.register_host_fn(MODULE, "searchParamsNew", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let init = args.first().cloned().unwrap_or(Value::Undefined);
        make_search_params_object(coerce_init_pairs(&init), None)
    }));

    vm.register_host_fn(MODULE, "searchParamsGet", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(Value::Object(obj)) = args.first() else {
            return Value::Null;
        };
        let key = args.get(1).map(|value| format!("{}", value)).unwrap_or_default();
        pairs_from_params(obj)
            .into_iter()
            .find_map(|(name, value)| (name == key).then(|| str_value(value)))
            .unwrap_or(Value::Null)
    }));

    vm.register_host_fn(MODULE, "searchParamsGetAll", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(Value::Object(obj)) = args.first() else {
            return params_array(Vec::new());
        };
        let key = args.get(1).map(|value| format!("{}", value)).unwrap_or_default();
        let values = pairs_from_params(obj)
            .into_iter()
            .filter_map(|(name, value)| (name == key).then(|| str_value(value)))
            .collect();
        params_array(values)
    }));

    vm.register_host_fn(MODULE, "searchParamsHas", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(Value::Object(obj)) = args.first() else {
            return Value::Bool(false);
        };
        let key = args.get(1).map(|value| format!("{}", value)).unwrap_or_default();
        Value::Bool(pairs_from_params(obj).into_iter().any(|(name, _)| name == key))
    }));

    vm.register_host_fn(MODULE, "searchParamsToString", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(Value::Object(obj)) = args.first() else {
            return str_value("");
        };
        str_value(serialize_pairs(&pairs_from_params(obj)))
    }));

    vm.register_host_fn(MODULE, "searchParamsAppend", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(Value::Object(obj)) = args.first() else {
            return Value::Undefined;
        };
        let key = args.get(1).map(|value| format!("{}", value)).unwrap_or_default();
        let value = args.get(2).map(|value| format!("{}", value)).unwrap_or_default();
        let mut pairs = pairs_from_params(obj);
        pairs.push((key, value));
        sync_pairs_object(obj, &pairs);
        update_owner_from_params(obj);
        Value::Undefined
    }));

    vm.register_host_fn(MODULE, "searchParamsSet", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(Value::Object(obj)) = args.first() else {
            return Value::Undefined;
        };
        let key = args.get(1).map(|value| format!("{}", value)).unwrap_or_default();
        let value = args.get(2).map(|value| format!("{}", value)).unwrap_or_default();
        let mut pairs = pairs_from_params(obj);
        if let Some(first_index) = pairs.iter().position(|(name, _)| *name == key) {
            pairs[first_index].1 = value.clone();
            let mut index = first_index + 1;
            while index < pairs.len() {
                if pairs[index].0 == key {
                    pairs.remove(index);
                } else {
                    index += 1;
                }
            }
        } else {
            pairs.push((key, value));
        }
        sync_pairs_object(obj, &pairs);
        update_owner_from_params(obj);
        Value::Undefined
    }));

    vm.register_host_fn(MODULE, "searchParamsDelete", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(Value::Object(obj)) = args.first() else {
            return Value::Undefined;
        };
        let key = args.get(1).map(|value| format!("{}", value)).unwrap_or_default();
        let pairs: Vec<(String, String)> = pairs_from_params(obj)
            .into_iter()
            .filter(|(name, _)| *name != key)
            .collect();
        sync_pairs_object(obj, &pairs);
        update_owner_from_params(obj);
        Value::Undefined
    }));

    vm.register_host_fn(MODULE, "searchParamsSort", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(Value::Object(obj)) = args.first() else {
            return Value::Undefined;
        };
        let mut pairs = pairs_from_params(obj);
        pairs.sort_by(|left, right| left.0.cmp(&right.0));
        sync_pairs_object(obj, &pairs);
        update_owner_from_params(obj);
        Value::Undefined
    }));

    vm.register_host_fn(MODULE, "searchParamsKeys", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(Value::Object(obj)) = args.first() else {
            return params_array(Vec::new());
        };
        params_array(
            pairs_from_params(obj)
                .into_iter()
                .map(|(key, _)| str_value(key))
                .collect(),
        )
    }));

    vm.register_host_fn(MODULE, "searchParamsValues", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(Value::Object(obj)) = args.first() else {
            return params_array(Vec::new());
        };
        params_array(
            pairs_from_params(obj)
                .into_iter()
                .map(|(_, value)| str_value(value))
                .collect(),
        )
    }));

    vm.register_host_fn(MODULE, "searchParamsEntries", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(Value::Object(obj)) = args.first() else {
            return params_array(Vec::new());
        };
        params_array(
            pairs_from_params(obj)
                .into_iter()
                .map(|(key, value)| pair_value(&key, &value))
                .collect(),
        )
    }));

    vm.register_host_fn(MODULE, "searchParamsIterator", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(Value::Object(obj)) = args.first() else {
            return make_bound_array_iterator(Vec::new());
        };
        make_bound_array_iterator(
            pairs_from_params(obj)
                .into_iter()
                .map(|(key, value)| pair_value(&key, &value))
                .collect(),
        )
    }));

    vm.register_host_fn(MODULE, "searchParamsForEach", Box::new(|ctx: &mut HostContext, args: &[Value]| {
        let Some(Value::Object(obj)) = args.first() else {
            return Value::Undefined;
        };
        let callback = args.get(1).cloned().unwrap_or(Value::Undefined);
        let this_arg = args.get(2).cloned();
        let receiver = Value::Object(obj.clone());
        let saved_this = this_arg.as_ref().map(|_| ctx.current_js_this());
        for (key, value) in pairs_from_params(obj) {
            let value_arg = str_value(value);
            let key_arg = str_value(key);
            if let Some(this_arg) = this_arg.clone() {
                ctx.set_js_this(this_arg);
                ctx.invoke(&callback, &[value_arg, key_arg, receiver.clone()]);
                if let Some(saved_this) = saved_this.clone() {
                    ctx.set_js_this(saved_this);
                }
            } else {
                ctx.invoke(&callback, &[value_arg, key_arg, receiver.clone()]);
            }
        }
        Value::Undefined
    }));

    let idx = |name: &str| {
        *vm.host_registry
            .get(&(MODULE.to_string(), name.to_string()))
            .expect("web:url host fn missing")
    };
    let _ = HOST_FNS.set(HostFns {
        url_to_string: idx("urlToString"),
        url_to_json: idx("urlToJSON"),
        set_href: idx("urlSetHref"),
        set_protocol: idx("urlSetProtocol"),
        set_username: idx("urlSetUsername"),
        set_password: idx("urlSetPassword"),
        set_host: idx("urlSetHost"),
        set_hostname: idx("urlSetHostname"),
        set_port: idx("urlSetPort"),
        set_pathname: idx("urlSetPathname"),
        set_search: idx("urlSetSearch"),
        set_hash: idx("urlSetHash"),
        params_get: idx("searchParamsGet"),
        params_get_all: idx("searchParamsGetAll"),
        params_has: idx("searchParamsHas"),
        params_to_string: idx("searchParamsToString"),
        params_append: idx("searchParamsAppend"),
        params_set: idx("searchParamsSet"),
        params_delete: idx("searchParamsDelete"),
        params_sort: idx("searchParamsSort"),
        params_keys: idx("searchParamsKeys"),
        params_values: idx("searchParamsValues"),
        params_entries: idx("searchParamsEntries"),
        params_iterator: idx("searchParamsIterator"),
        params_for_each: idx("searchParamsForEach"),
    });
}
