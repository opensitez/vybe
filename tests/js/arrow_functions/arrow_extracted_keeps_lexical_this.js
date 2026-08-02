// vybe-test: js/arrow_functions/arrow_extracted_keeps_lexical_this
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

const holder={v:7,make(){return ()=>this.v;}}; const ext=holder.make(); __check(__line(ext()), "7");
