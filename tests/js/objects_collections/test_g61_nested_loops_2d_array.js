// vybe-test: js/objects_collections/test_g61_nested_loops_2d_array
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

let grid = [[1, 2], [3, 4], [5, 6]];
        let sum = 0;
        let i = 0;
        while (i < grid.length) {
            let j = 0;
            while (j < grid[i].length) {
                sum = sum + grid[i][j];
                j = j + 1;
            }
            i = i + 1;
        }
        console.log(sum);
