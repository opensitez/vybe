// vybe-test: js/array_find_last_and_find_last_index/test_js_array_find_last_mutation_during_iteration
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

const arr = [1, 2, 3];
const visited = [];
arr.findLast((val) => {
    visited.push(val);
    if (val === 3) arr.pop(); // Pop element 3 during iteration
    return false;
});
__check(__line(visited.join(",")), "3,2,1");
