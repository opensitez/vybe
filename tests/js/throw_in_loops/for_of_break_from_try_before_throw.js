// vybe-test: js/throw_in_loops/for_of_break_from_try_before_throw
// origin: languages/js/tests/js/test_throw_in_loops.rs

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

let o=[];for(const x of [1,2,3]){try{if(x===2)break;o.push(x);}catch{o.push("e");}}console.log(o.join(","));
