// vybe-test: js/object_introspection/object_keys_own_enumerable_only
// origin: languages/js/tests/js/test_object_introspection.rs

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

const proto = { inherited: 0 };
const obj = Object.create(proto);
obj.a = 1;
obj.b = 2;
Object.defineProperty(obj, "hidden", { value: 3, enumerable: false });
const keys = Object.keys(obj);
__check(__line(keys.join(",")), "a,b");
__check(__line(keys.includes("inherited")), "false");
__check(__line(keys.includes("hidden")), "false");
