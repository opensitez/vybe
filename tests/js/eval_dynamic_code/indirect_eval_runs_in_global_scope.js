// vybe-test: js/eval_dynamic_code/indirect_eval_runs_in_global_scope
// origin: languages/js/tests/js/test_eval_dynamic_code.rs

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

const x = "local";
globalThis.x = "global";
const indirectEval = eval;
const result = indirectEval("x");
__check(__line(result), "global");
