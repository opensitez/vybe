// vybe-test: js/array_find_last_and_find_last_index/test_js_array_find_last_object_elements
// origin: languages/js/tests/js/test_js_array_find_last_and_find_last_index.rs

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

const users = [
    { id: 1, active: true },
    { id: 2, active: false },
    { id: 3, active: true }
];
const lastActive = users.findLast(u => u.active);
__check(__line(lastActive.id), "3");
