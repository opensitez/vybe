// vybe-test: js/interop/test_f59_closure_inside_class_method
// origin: languages/js/tests/js/js_interop_test.rs

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

class Processor {
            constructor(data) { this.data = data; }
            processAll() {
                let results = [];
                let self = this;
                this.data.forEach(x => {
                    results.push(x * 2);
                });
                return results;
            }
        }
        let p = new Processor([1, 2, 3]);
        console.log(p.processAll());
