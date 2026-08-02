// vybe-test: js/proxy_invariant_violations/proxy_delete_trap_can_block_delete
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

const t={a:1}; const p=new Proxy(t,{deleteProperty(){return false;}}); delete p.a; __check(__line(t.a), "1");
