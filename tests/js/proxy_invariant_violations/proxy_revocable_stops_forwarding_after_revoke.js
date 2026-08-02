// vybe-test: js/proxy_invariant_violations/proxy_revocable_stops_forwarding_after_revoke
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

const {proxy,revoke}=Proxy.revocable({x:1},{}); revoke(); try{console.log(proxy.x);}catch(e){console.log(e instanceof TypeError);}
