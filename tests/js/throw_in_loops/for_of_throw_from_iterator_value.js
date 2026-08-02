// vybe-test: js/throw_in_loops/for_of_throw_from_iterator_value
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

let o=[];try{for(const x of [1,2,3]){if(x===2)throw new Error("v"+x);o.push(x);}}catch(e){o.push(e.message);}console.log(o.join(","));
