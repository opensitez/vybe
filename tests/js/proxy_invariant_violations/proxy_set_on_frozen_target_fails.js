// vybe-test: js/proxy_invariant_violations/proxy_set_on_frozen_target_fails
// origin: languages/js/tests/js/test_proxy_invariant_violations.rs

function __line(...args) {
    // console.log joins its arguments with a single space. String() is the
    // coercion Vybe's logging host applies to each one.
    return args.map(String).join(" ");
}

function __check(got, want) {
    if (got !== want) {
        console.log("FAIL: want [" + want + "] got [" + got + "]");
        throw new Error("assertion failed");
    }
}

const t=Object.freeze({}); const p=new Proxy(t,{}); __check(__line(Reflect.set(p,"x",1)), "false");
