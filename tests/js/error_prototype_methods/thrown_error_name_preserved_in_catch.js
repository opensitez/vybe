// vybe-test: js/error_prototype_methods/thrown_error_name_preserved_in_catch
// origin: languages/js/tests/js/test_error_prototype_methods.rs

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

try{throw new URIError("bad");}catch(e){__check(__line(e.name), "URIError");}
