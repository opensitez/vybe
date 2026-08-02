// vybe-test: js/throw_in_loops/for_in_throw_custom_object_key
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

let o=[];const obj={a:1,b:2};try{for(const k in obj){if(k==="b")throw new Error(k);o.push(k);}}catch(e){o.push(e.message);}console.log(o.join(","));
