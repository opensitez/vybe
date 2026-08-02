// vybe-test: js/symbol_wellknown/symbol_isconcatspreadable_array_like
// origin: languages/js/tests/js/test_symbol_wellknown.rs

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

const arrayLike = { 0: "a", 1: "b", length: 2, [Symbol.isConcatSpreadable]: true };
const result = ["x"].concat(arrayLike);
__check(__line(result.join(",")), "x,a,b");
