// vybe-test: js/throw_in_loops/for_of_iterator_return_value_typeerror
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

let o=[];try{const bad={
            [Symbol.iterator]() {
                return {
                    next() {
                        if (this._n === undefined) this._n = 0;
                        this._n++;
                        return this._n === 1 ? { value: 1, done: false } : 0;
                    }
                };
            }
        };for (const x of bad) o.push(x);}
        catch (e) { o.push(e instanceof TypeError ? "TypeError" : "Other"); }
        console.log(o.join(","));
