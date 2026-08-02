// vybe-test: js/arrow_functions/nested_arrows_share_lexical_this
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

const deep={v:3,run(){return (()=>(()=>this.v)())();}}; __check(__line(deep.run()), "3");
