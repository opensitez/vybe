// vybe-test: js/property_enumeration/integer_indices_come_before_string_keys
// origin: languages/js/tests/js/test_property_enumeration.rs

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

const obj = {};
obj.b = 1;
obj[2] = 2;
obj.a = 3;
obj[1] = 4;
obj[0] = 5;
const keys = Object.keys(obj);
const intKeys = keys.filter(k => /^\d+$/.test(k)).sort((a,b) => +a - +b);
const strKeys = keys.filter(k => !/^\d+$/.test(k));
__check(__line([...intKeys, ...strKeys].join(",")), "0,1,2,b,a");
