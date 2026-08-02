// vybe-test: js/immutable_patterns/persistent_stack
// origin: languages/js/tests/js/test_immutable_patterns.rs

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

class Stack {
    constructor(head = null, tail = null) { this.head = head; this.tail = tail; }
    push(val) { return new Stack(val, this); }
    pop() { return this.tail; }
    peek() { return this.head; }
    get isEmpty() { return this.head === null; }
}
const s0 = new Stack();
const s1 = s0.push(1);
const s2 = s1.push(2);
const s3 = s2.push(3);
__check(__line(s3.peek()), "3");
__check(__line(s2.peek()), "2");  // s2 still intact
__check(__line(s3.pop().peek()), "2");
