// vybe-test: js/proxy_invariant_violations/proxy_apply_trap_on_function_target
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

const fn=function(a,b){return a+b;}; const p=new Proxy(fn,{apply(t,_,args){return args[0]*args[1];}}); __check(__line(p(3,4)), "12");
