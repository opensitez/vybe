// vybe-test: js/json_semantics_matrix/json_stringify_custom_tojson_and_parse_reviver_compose
// origin: languages/js/tests/js/test_json_semantics_matrix.rs

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

const src = {
    item: {
        value: 4,
        toJSON() {
            return { value: this.value * 2 };
        }
    }
};
const back = JSON.parse(JSON.stringify(src), (key, value) => {
    return key === "value" ? value / 2 : value;
});
__check(__line(back.item.value), "4");
