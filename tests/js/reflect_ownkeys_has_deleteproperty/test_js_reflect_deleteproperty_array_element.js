// vybe-test: js/reflect_ownkeys_has_deleteproperty/test_js_reflect_deleteproperty_array_element
// origin: languages/js/tests/js/test_js_reflect_ownkeys_has_deleteproperty.rs

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

const arr = [10, 20, 30];
const res = Reflect.deleteProperty(arr, 1);
__check(__line(res + "|len=" + arr.length + "|hasIndex1=" + (1 in arr)), "true|len=3|hasIndex1=false");
