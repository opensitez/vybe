// vybe-test: js/array_group_by_and_group_by_to_map/test_js_object_groupby_callback_arguments
// origin: languages/js/tests/js/test_js_array_group_by_and_group_by_to_map.rs

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

const arr = ["a", "b"];
const log = [];
Object.groupBy(arr, (val, index) => {
    log.push(`${val}:${index}`);
    return "group";
});
__check(__line(log.join("|")), "a:0|b:1");
