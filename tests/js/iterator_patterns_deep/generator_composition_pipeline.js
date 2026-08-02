// vybe-test: js/iterator_patterns_deep/generator_composition_pipeline
// origin: languages/js/tests/js/test_iterator_patterns_deep.rs

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

function* map(iter, fn) { for (const v of iter) yield fn(v); }
function* filter(iter, pred) { for (const v of iter) if (pred(v)) yield v; }
function* take(iter, n) { let i=0; for (const v of iter) { if (i++>=n) break; yield v; } }
function* counter(start=0) { while(true) yield start++; }

const result = [...take(filter(map(counter(), x=>x*x), x=>x%2===0), 4)];
console.log(result.join(","));
