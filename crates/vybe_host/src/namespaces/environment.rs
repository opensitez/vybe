use super::*;

pub fn register(vm: &mut VM) {
    // JS `env.*` ambient compatibility namespace — backed by wasi:cli/environment.
    let js_env = ensure_namespace(vm, &["env"]);
    set_prop(&js_env, "args", host_fn_ref(vm, "wasi:cli/environment", "get-arguments"));
    set_prop(&js_env, "cwd", host_fn_ref(vm, "wasi:cli/environment", "initial-cwd"));
    set_prop(&js_env, "getEnv", host_fn_ref(vm, "wasi:cli/environment", "get-environment"));

    // JS `random.*` ambient compatibility namespace.
    let random = ensure_namespace(vm, &["random"]);
    set_prop(&random, "get-random-bytes", host_fn_ref(vm, "wasi:random/random", "get-random-bytes"));
    set_prop(&random, "get-random-u64", host_fn_ref(vm, "wasi:random/random", "get-random-u64"));
    set_prop(&random, "random", host_fn_ref(vm, "wasi:random/random", "random"));
    set_prop(&random, "randomInt", host_fn_ref(vm, "wasi:random/random", "randomInt"));
    set_prop(&random, "uuid", host_fn_ref(vm, "wasi:random/random", "uuid"));

    // JS `http.*` ambient namespace (wasi:http shim).
    let http = ensure_namespace(vm, &["http"]);
    set_prop(&http, "fetch", host_fn_ref(vm, "wasi:http", "fetch"));
    set_prop(&http, "get", host_fn_ref(vm, "wasi:http", "get"));
    set_prop(&http, "post", host_fn_ref(vm, "wasi:http", "post"));
}
