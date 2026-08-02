// vybe-test: js/private_brand_errors/instance_cannot_read_private_static_field
// origin: languages/js/tests/js/test_private_brand_errors.rs

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

class C{static #s=1;} const c=new C(); try{console.log(c.#s);}catch(e){console.log(e instanceof TypeError);}
