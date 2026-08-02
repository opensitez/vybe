// vybe-test: js/private_brand_errors/destructuring_private_field_outside_class_throws
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

class C{ #x=1; static read(o){ return o.#x; } } try{ C.read({}); }catch(e){ __check(__line(e instanceof TypeError), "true"); }
