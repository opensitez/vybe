// vybe-test: js/proxy_reflect/proxy_array_length_enforcement
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

function createBoundedArray(max) {
    return new Proxy([], {
        set(target, prop, value) {
            if (prop === "length" && value > max) {
                throw new RangeError("array too large");
            }
            target[prop] = value;
            return true;
        }
    });
}
const arr = createBoundedArray(3);
arr[0] = 1;
arr[1] = 2;
arr[2] = 3;
__check(__line(arr[0]), "1");
__check(__line(arr[2]), "3");
