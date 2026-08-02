// vybe-test: js/proxy_invariant_violations/proxy_set_prototype_of_forwards
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

const t={}; const np={}; const p=new Proxy(t,{}); Object.setPrototypeOf(p,np); __check(__line(Object.getPrototypeOf(t)===np), "true");
