// vybe-test: js/try_catch_nested/nested_error_object_identity_preserved
// origin: languages/js/tests/js/test_try_catch_nested.rs

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

let obj={id:7};
let same=false;
try{try{throw obj;}catch(e){same=(e===obj);}}
catch{same=false;}
__check(__line(same), "true");
