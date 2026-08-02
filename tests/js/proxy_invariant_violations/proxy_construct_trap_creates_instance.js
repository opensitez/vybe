// vybe-test: js/proxy_invariant_violations/proxy_construct_trap_creates_instance
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

class C{constructor(v){this.v=v;}} const p=new Proxy(C,{construct(t,args){return new t(args[0]*2);}}); __check(__line(new p(3).v), "6");
