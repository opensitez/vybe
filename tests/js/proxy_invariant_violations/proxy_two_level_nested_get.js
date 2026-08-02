// vybe-test: js/proxy_invariant_violations/proxy_two_level_nested_get
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

const inner=new Proxy({v:1},{}); const outer=new Proxy({inner},{get(t,k){return t[k];}}); __check(__line(outer.inner.v), "1");
