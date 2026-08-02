// vybe-test: js/array_iteration_methods/reduceright_processes_right_to_left
// origin: languages/js/tests/js/test_array_iteration_methods.rs

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

const result = [[1,2],[3,4],[5,6]].reduceRight((acc, x) => acc.concat(x), []);
__check(__line(result.join(",")), "5,6,3,4,1,2");
