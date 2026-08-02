// vybe-test: js/private_brand_errors/private_in_on_primitive_throws
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

class C{#x=1; static check(v){return #x in v;}} try{C.check(1);}catch(e){__check(__line(e instanceof TypeError), "true");}
