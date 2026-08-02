// vybe-test: js/map_set_prototype_methods/map_for_each_insertion_order
// origin: languages/js/tests/js/test_map_set_prototype_methods.rs

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

const m=new Map([["a",1],["b",2]]); const o=[]; m.forEach((v,k)=>o.push(k+":"+v)); console.log(o.join(","));
