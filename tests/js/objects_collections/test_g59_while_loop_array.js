// vybe-test: js/objects_collections/test_g59_while_loop_array
// origin: languages/js/tests/js/js_objects_collections_test.rs

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

let arr = [5, 10, 15, 20];
        let i = 0;
        let sum = 0;
        while (i < arr.length) {
            sum = sum + arr[i];
            i = i + 1;
        }
        console.log(sum);
