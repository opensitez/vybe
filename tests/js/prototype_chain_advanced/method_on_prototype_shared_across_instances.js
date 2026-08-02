// vybe-test: js/prototype_chain_advanced/method_on_prototype_shared_across_instances
// origin: languages/js/tests/js/test_prototype_chain_advanced.rs

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

function Counter() { this.count = 0; }
Counter.prototype.increment = function() { this.count++; };

const c1 = new Counter();
const c2 = new Counter();

c1.increment(); c1.increment();
c2.increment();

__check(__line(c1.count), "2");
__check(__line(c2.count), "1");
// Shared method
__check(__line(c1.increment === c2.increment), "true");
