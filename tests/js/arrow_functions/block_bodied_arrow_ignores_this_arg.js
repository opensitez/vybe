// vybe-test: js/arrow_functions/block_bodied_arrow_ignores_this_arg
// origin: languages/js/tests/js/test_arrow_functions.rs

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

const blk=function(){ return (() => { return this; }); }.call({tag:"outer"})(); __check(__line(blk && blk.tag), "outer");
