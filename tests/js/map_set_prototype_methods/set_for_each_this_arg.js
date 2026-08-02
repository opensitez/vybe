// vybe-test: js/map_set_prototype_methods/set_for_each_this_arg
// origin: languages/js/tests/js/test_map_set_prototype_methods.rs

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

const s=new Set([1]); const ctx={n:0}; s.forEach(function(v){this.n+=v;},ctx); console.log(ctx.n);
