// vybe-test: js/object_immutability/prevent_extensions_on_array_rejects_length_growth
// origin: languages/js/tests/js/test_object_immutability.rs

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

const arr = [1, 2, 3];
Object.preventExtensions(arr);
let threw = false;
try {
    arr.push(4);
} catch {
    threw = true;
}
arr[4] = 5;
__check(__line(threw), "true");
__check(__line(arr.length), "3");
__check(__line(arr[3]), "undefined");
__check(__line(arr[4]), "undefined");
