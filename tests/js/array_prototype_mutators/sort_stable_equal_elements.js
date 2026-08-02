// vybe-test: js/array_prototype_mutators/sort_stable_equal_elements
// origin: languages/js/tests/js/test_array_prototype_mutators.rs

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

const a=[{v:1},{v:2},{v:1}]; a.sort((a,b)=>a.v-b.v); __check(__line(a.map(x=>x.v).join(",")), "1,1,2");
