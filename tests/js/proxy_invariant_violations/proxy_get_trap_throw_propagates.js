// vybe-test: js/proxy_invariant_violations/proxy_get_trap_throw_propagates
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

const p=new Proxy({},{get(){throw new Error("trap");}}); try{console.log(p.a);}catch(e){console.log(e.message);}
