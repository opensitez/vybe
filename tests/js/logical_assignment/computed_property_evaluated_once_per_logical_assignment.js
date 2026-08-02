// vybe-test: js/logical_assignment/computed_property_evaluated_once_per_logical_assignment
// origin: languages/js/tests/js/test_logical_assignment.rs

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

let compute_count = 0;
const key = () => {
    compute_count++;
    return "flag";
};
const obj = { flag: 0 };

obj[key()] ||= 1;
obj[key()] ||= 2;
obj[key()] &&= 3;

__check(__line(obj.flag), "3");
__check(__line(compute_count), "6");
