// vybe-test: js/property_ordering/get_own_property_names_includes_non_enumerable
// origin: languages/js/tests/js/test_property_ordering.rs

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

const obj = { a: 1 };
Object.defineProperty(obj, "hidden", { value: 2, enumerable: false });
const names = Object.getOwnPropertyNames(obj);
__check(__line(names.includes("a")), "true");
__check(__line(names.includes("hidden")), "true");
