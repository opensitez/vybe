// vybe-test: js/array_prototype_mutators/unshift_on_long_array
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

const a=Array.from({length:100},(_,i)=>i); a.unshift(-1); __check(__line(a[0]), "-1"); __check(__line(a.length), "101");
