// vybe-test: js/array_prototype_mutators/sort_comparator_not_callable_throws
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

try{[1,2].sort({}); console.log("ok");}catch(e){console.log(e instanceof TypeError);}
