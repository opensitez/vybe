// vybe-test: js/proxy_reflect/reflect_is_extensible
// origin: languages/js/tests/js/test_proxy_reflect.rs

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

const obj = {};
__check(__line(Reflect.isExtensible(obj)), "true");
// A fresh object is extensible
__check(__line(typeof obj === "object"), "true");
