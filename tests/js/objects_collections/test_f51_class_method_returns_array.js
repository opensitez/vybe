// vybe-test: js/objects_collections/test_f51_class_method_returns_array
// origin: languages/js/tests/js/js_objects_collections_test.rs

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

class NumGen {
            constructor(n) { this.n = n; }
            generate() {
                let arr = [];
                let i = 0;
                while (i < this.n) { arr.push(i); i = i + 1; }
                return arr;
            }
        }
        let g = new NumGen(4);
        console.log(g.generate().join(","));
