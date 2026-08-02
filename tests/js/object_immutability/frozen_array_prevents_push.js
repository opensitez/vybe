// vybe-test: js/object_immutability/frozen_array_prevents_push
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

const arr = Object.freeze([1, 2, 3]);
try { arr.push(4); } catch {}
__check(__line(arr.length), "3");
__check(__line(arr[3]), "undefined");
