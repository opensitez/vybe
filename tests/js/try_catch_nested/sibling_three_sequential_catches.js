// vybe-test: js/try_catch_nested/sibling_three_sequential_catches
// origin: languages/js/tests/js/test_try_catch_nested.rs

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

let o=[];
try{throw 1;}catch{o.push("1");}
try{throw 2;}catch{o.push("2");}
try{throw 3;}catch{o.push("3");}
__check(__line(o.join(",")), "1,2,3");
