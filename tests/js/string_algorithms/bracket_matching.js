// vybe-test: js/string_algorithms/bracket_matching
// origin: languages/js/tests/js/test_string_algorithms.rs

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

function isValid(s) {
    const stack = [];
    const map = { ')':'(', ']':'[', '}':'{' };
    for (const c of s) {
        if ("([{".includes(c)) stack.push(c);
        else if (stack.pop() !== map[c]) return false;
    }
    return stack.length === 0;
}
console.log(isValid("()[]{}"));
console.log(isValid("([)]"));
console.log(isValid("{[]}"));
