// vybe-test: js/array_sort_to_sorted_comparator_stability/test_js_array_sort_object_by_property
// origin: languages/js/tests/js/test_js_array_sort_to_sorted_comparator_stability.rs

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

const users = [{ age: 30 }, { age: 20 }, { age: 25 }];
users.sort((a, b) => a.age - b.age);
__check(__line(users.map(u => u.age).join(",")), "20,25,30");
