// vybe-test: js/object_prototype_methods/hasownproperty_prototype_method_not_own
// origin: languages/js/tests/js/test_object_prototype_methods.rs

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

function C(){} C.prototype.m=function(){}; const c=new C(); __check(__line(c.hasOwnProperty("m")), "false");
