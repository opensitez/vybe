// vybe-test: js/misc_es_features/array_groupby_with_reduce
// origin: languages/js/tests/js/test_misc_es_features.rs

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

const items = ["apple", "banana", "cherry", "avocado"];
const grouped = items.reduce((acc, word) => {
  const key = word[0];
  (acc[key] = acc[key] || []).push(word);
  return acc;
}, {});
__check(__line(grouped["a"].length), "2");
__check(__line(grouped["b"][0]), "banana");
