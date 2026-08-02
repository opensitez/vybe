// vybe-test: js/block_scoping_let_const_in_loops_and_blocks/test_js_for_of_loop_const_binding
// origin: languages/js/tests/js/test_js_block_scoping_let_const_in_loops_and_blocks.rs

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

const arr = [10, 20];
const vals = [];
for (const x of arr) {
    vals.push(x * 2);
}
console.log(vals.join(","));
