// vybe-test: js/structured_clone_advanced/structuredclone_regexp
// origin: languages/js/tests/js/test_structured_clone_advanced.rs

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

const re = /hello/gi;
const clone = structuredClone(re);
__check(__line(clone instanceof RegExp), "true");
__check(__line(clone.flags.includes("g")), "true");
