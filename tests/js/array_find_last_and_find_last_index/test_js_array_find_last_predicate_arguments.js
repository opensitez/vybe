// vybe-test: js/array_find_last_and_find_last_index/test_js_array_find_last_predicate_arguments
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

const arr = ["a", "b", "c"];
const log = [];
arr.findLast((val, index, array) => {
    log.push(`${val}:${index}:${array.length}`);
    return false;
});
__check(__line(log.join("|")), "c:2:3|b:1:3|a:0:3");
