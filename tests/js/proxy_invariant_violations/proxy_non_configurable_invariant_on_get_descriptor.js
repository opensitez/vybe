// vybe-test: js/proxy_invariant_violations/proxy_non_configurable_invariant_on_get_descriptor
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

const t={}; Object.defineProperty(t,"x",{value:1,configurable:false}); const p=new Proxy(t,{getOwnPropertyDescriptor(){return {value:2,configurable:false,enumerable:true,writable:true};}}); try{Object.getOwnPropertyDescriptor(p,"x");}catch(e){__check(__line(e instanceof TypeError), "true");}
