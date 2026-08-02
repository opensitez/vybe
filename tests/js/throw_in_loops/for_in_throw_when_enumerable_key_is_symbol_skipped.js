// vybe-test: js/throw_in_loops/for_in_throw_when_enumerable_key_is_symbol_skipped
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

let o=[];const s=Symbol("s");const obj={a:1};obj[s]=2;try{for(const k in obj){if(k==="a")throw new Error("a");o.push(k);}}catch(e){o.push(e.message);}console.log(o.join(","));
