// vybe-test: js/symbol_protocols/symbol_concat_spreadable
// origin: languages/js/tests/js/test_symbol_protocols.rs

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

const arrayLike = { 0: "a", 1: "b", 2: "c", length: 3, [Symbol.isConcatSpreadable]: true };
const result = ["x"].concat(arrayLike);
__check(__line(result.join(",")), "x,a,b,c");
const notSpreadable = [1, 2];
notSpreadable[Symbol.isConcatSpreadable] = false;
const result2 = ["y"].concat(notSpreadable);
__check(__line(result2.length), "2");
