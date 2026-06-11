use super::*;
use vybe_bytecode::HostContext;

pub fn register(vm: &mut VM) {
    vm.register_host_fn(
        "vybe:compat/clock",
        "now",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            let ms = std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .map(|dur| dur.as_millis() as f64)
                .unwrap_or(0.0);
            Value::F64(ms)
        }),
    );

    vm.register_host_fn(
        "vybe:compat/env",
        "platform",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            let platform = match std::env::consts::OS {
                "macos" => "macos",
                "windows" => "windows",
                _ => "linux",
            };
            Value::String(Arc::from(platform))
        }),
    );

    vm.register_host_fn(
        "vybe:compat/http",
        "fetch",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let url = args.first().map(|v| format!("{}", v)).unwrap_or_default();
            let ok = url
                .strip_prefix("http://")
                .and_then(|rest| rest.split('/').next())
                .and_then(|authority| {
                    let (host, port) = authority
                        .rsplit_once(':')
                        .map(|(host, port)| (host, port.parse::<u16>().ok()))
                        .unwrap_or((authority, Some(80)));
                    port.map(|port| (host.to_string(), port))
                })
                .is_some_and(|(host, port)| {
                    std::net::TcpStream::connect((host.as_str(), port)).is_ok()
                });
            let mut response = Object::new();
            response.properties.insert("ok".into(), Value::Bool(ok));
            response
                .properties
                .insert("status".into(), Value::F64(if ok { 200.0 } else { 0.0 }));
            Value::Object(Arc::new(Mutex::new(response)))
        }),
    );

    let clock = ensure_namespace(vm, &["clock"]);
    set_prop(&clock, "now", host_fn_ref(vm, "vybe:compat/clock", "now"));
    set_prop(
        &clock,
        "toISOString",
        host_fn_ref(vm, "ecma:date", "toISOString"),
    );

    // JS `env.*` ambient compatibility namespace — backed by wasi:cli/environment.
    let js_env = ensure_namespace(vm, &["env"]);
    set_prop(
        &js_env,
        "args",
        host_fn_ref(vm, "wasi:cli/environment", "get-arguments"),
    );
    set_prop(
        &js_env,
        "cwd",
        host_fn_ref(vm, "wasi:cli/environment", "initial-cwd"),
    );
    set_prop(
        &js_env,
        "getEnv",
        host_fn_ref(vm, "wasi:cli/environment", "get-environment"),
    );
    set_prop(
        &js_env,
        "platform",
        host_fn_ref(vm, "vybe:compat/env", "platform"),
    );

    // JS `random.*` ambient compatibility namespace.
    let random = ensure_namespace(vm, &["random"]);
    set_prop(
        &random,
        "get-random-bytes",
        host_fn_ref(vm, "wasi:random/random", "get-random-bytes"),
    );
    set_prop(
        &random,
        "get-random-u64",
        host_fn_ref(vm, "wasi:random/random", "get-random-u64"),
    );
    set_prop(
        &random,
        "random",
        host_fn_ref(vm, "wasi:random/random", "random"),
    );
    set_prop(
        &random,
        "randomInt",
        host_fn_ref(vm, "wasi:random/random", "randomInt"),
    );
    set_prop(
        &random,
        "uuid",
        host_fn_ref(vm, "wasi:random/random", "uuid"),
    );

    // JS `http.*` ambient namespace (wasi:http shim).
    let http = ensure_namespace(vm, &["http"]);
    set_prop(&http, "fetch", host_fn_ref(vm, "vybe:compat/http", "fetch"));
    set_prop(&http, "get", host_fn_ref(vm, "wasi:http", "get"));
    set_prop(&http, "post", host_fn_ref(vm, "wasi:http", "post"));
}
