/// JavaScript advanced string and array methods not covered elsewhere:
/// String: replaceAll, matchAll, trimStart/trimEnd, at(), raw, fromCharCode
/// Array: flatMap, at(), Array.of, from with mapper, reduceRight,
/// copyWithin, entries/keys/values
use super::helpers::run_js;

// ===================================================================
// STRING: AT()
// ===================================================================

#[test]
fn string_at_positive() {
    assert_eq!(
        run_js(
            r#"
let s = "Hello";
console.log(s.at(0));
console.log(s.at(1));
"#
        ),
        &["H", "e"]
    );
}

#[test]
fn string_at_negative() {
    assert_eq!(
        run_js(
            r#"
let s = "Hello";
console.log(s.at(-1));
console.log(s.at(-2));
"#
        ),
        &["o", "l"]
    );
}

// ===================================================================
// STRING: REPLACEALL
// ===================================================================

#[test]
fn string_replace_all() {
    assert_eq!(
        run_js(
            r#"
let s = "foo-bar-baz-foo";
console.log(s.replaceAll("foo", "qux"));
"#
        ),
        &["qux-bar-baz-qux"]
    );
}

#[test]
fn string_replace_all_empty() {
    assert_eq!(
        run_js(
            r#"
let s = "abc";
console.log(s.replaceAll("x", "y"));
"#
        ),
        &["abc"]
    );
}

// ===================================================================
// STRING: TRIMSTART / TRIMEND
// ===================================================================

#[test]
fn string_trim_start() {
    assert_eq!(
        run_js(
            r#"
console.log("   hello   ".trimStart());
"#
        ),
        &["hello   "]
    );
}

#[test]
fn string_trim_end() {
    assert_eq!(
        run_js(
            r#"
console.log("   hello   ".trimEnd());
"#
        ),
        &["   hello"]
    );
}

// ===================================================================
// STRING: SEARCH / MATCH PATTERNS
// ===================================================================

#[test]
fn string_search() {
    assert_eq!(
        run_js(
            r#"
let s = "Hello World";
console.log(s.search("World"));
console.log(s.search("xyz"));
"#
        ),
        &["6", "-1"]
    );
}

#[test]
fn string_match_basic() {
    assert_eq!(
        run_js(
            r#"
let s = "The year is 2024, not 2023";
let matches = s.match(/\d+/g);
console.log(matches.join(","));
"#
        ),
        &["2024,2023"]
    );
}

#[test]
fn string_match_groups() {
    assert_eq!(
        run_js(
            r#"
let s = "2024-01-15";
let m = s.match(/(\d{4})-(\d{2})-(\d{2})/);
console.log(m[1]);
console.log(m[2]);
console.log(m[3]);
"#
        ),
        &["2024", "01", "15"]
    );
}

// ===================================================================
// STRING: SLICE VS SUBSTRING
// ===================================================================

#[test]
fn string_slice_negative() {
    assert_eq!(
        run_js(
            r#"
let s = "Hello World";
console.log(s.slice(-5));
console.log(s.slice(-5, -1));
"#
        ),
        &["World", "Worl"]
    );
}

#[test]
fn string_concat_method() {
    assert_eq!(
        run_js(
            r#"
let s = "Hello";
console.log(s.concat(" ", "World", "!"));
"#
        ),
        &["Hello World!"]
    );
}

// ===================================================================
// ARRAY: AT()
// ===================================================================

#[test]
fn array_at_positive() {
    assert_eq!(
        run_js(
            r#"
let arr = [10, 20, 30, 40, 50];
console.log(arr.at(0));
console.log(arr.at(2));
"#
        ),
        &["10", "30"]
    );
}

#[test]
fn array_at_negative() {
    assert_eq!(
        run_js(
            r#"
let arr = [10, 20, 30, 40, 50];
console.log(arr.at(-1));
console.log(arr.at(-2));
"#
        ),
        &["50", "40"]
    );
}

// ===================================================================
// ARRAY: FLATMAP
// ===================================================================

#[test]
fn array_flatmap() {
    assert_eq!(
        run_js(
            r#"
let arr = [1, 2, 3];
let result = arr.flatMap(x => [x, x * 2]);
console.log(result.join(","));
"#
        ),
        &["1,2,2,4,3,6"]
    );
}

#[test]
fn array_flatmap_filter() {
    assert_eq!(
        run_js(
            r#"
let arr = ["hello world", "foo bar"];
let words = arr.flatMap(s => s.split(" "));
console.log(words.join(","));
"#
        ),
        &["hello,world,foo,bar"]
    );
}

// ===================================================================
// ARRAY: OF / FROM WITH MAPPER
// ===================================================================

#[test]
fn array_of() {
    assert_eq!(
        run_js(
            r#"
let arr = Array.of(1, 2, 3);
console.log(arr.join(","));
console.log(arr.length);
"#
        ),
        &["1,2,3", "3"]
    );
}

#[test]
fn array_from_with_mapper() {
    assert_eq!(
        run_js(
            r#"
let arr = Array.from([1, 2, 3], x => x * x);
console.log(arr.join(","));
"#
        ),
        &["1,4,9"]
    );
}

#[test]
fn array_from_string() {
    assert_eq!(
        run_js(
            r#"
let arr = Array.from("hello");
console.log(arr.join(","));
"#
        ),
        &["h,e,l,l,o"]
    );
}

#[test]
fn array_from_length_object() {
    assert_eq!(
        run_js(
            r#"
let arr = Array.from({ length: 5 }, (_, i) => i * 2);
console.log(arr.join(","));
"#
        ),
        &["0,2,4,6,8"]
    );
}

// ===================================================================
// ARRAY: REDUCERIGHT
// ===================================================================

#[test]
fn array_reduce_right() {
    assert_eq!(
        run_js(
            r#"
let arr = ["a", "b", "c", "d"];
let result = arr.reduceRight((acc, val) => acc + val, "");
console.log(result);
"#
        ),
        &["dcba"]
    );
}

// ===================================================================
// ARRAY: SPLICE ADVANCED
// ===================================================================

#[test]
fn array_splice_remove_and_insert() {
    assert_eq!(
        run_js(
            r#"
let arr = [1, 2, 3, 4, 5];
let removed = arr.splice(1, 2, 10, 20, 30);
console.log(removed.join(","));
console.log(arr.join(","));
"#
        ),
        &["2,3", "1,10,20,30,4,5"]
    );
}

#[test]
fn array_splice_insert_only() {
    assert_eq!(
        run_js(
            r#"
let arr = [1, 2, 3];
arr.splice(1, 0, 99);
console.log(arr.join(","));
"#
        ),
        &["1,99,2,3"]
    );
}

// ===================================================================
// ARRAY: UNSHIFT
// ===================================================================

#[test]
fn array_unshift() {
    assert_eq!(
        run_js(
            r#"
let arr = [3, 4, 5];
arr.unshift(1, 2);
console.log(arr.join(","));
"#
        ),
        &["1,2,3,4,5"]
    );
}

// ===================================================================
// ARRAY: COPYWITHIN
// ===================================================================

#[test]
fn array_copywithin() {
    assert_eq!(
        run_js(
            r#"
let arr = [1, 2, 3, 4, 5];
arr.copyWithin(0, 3);
console.log(arr.join(","));
"#
        ),
        &["4,5,3,4,5"]
    );
}

// ===================================================================
// ARRAY: ENTRIES / KEYS / VALUES
// ===================================================================

#[test]
fn array_entries() {
    assert_eq!(
        run_js(
            r#"
let arr = ["a", "b", "c"];
for (let [i, v] of arr.entries()) {
    console.log(i + ":" + v);
}
"#
        ),
        &["0:a", "1:b", "2:c"]
    );
}

#[test]
fn array_keys() {
    assert_eq!(
        run_js(
            r#"
let arr = ["x", "y", "z"];
let keys = [...arr.keys()];
console.log(keys.join(","));
"#
        ),
        &["0,1,2"]
    );
}

#[test]
fn array_values() {
    assert_eq!(
        run_js(
            r#"
let arr = [10, 20, 30];
let vals = [...arr.values()];
console.log(vals.join(","));
"#
        ),
        &["10,20,30"]
    );
}

// ===================================================================
// ARRAY: SORTING PATTERNS
// ===================================================================

#[test]
fn array_sort_strings() {
    assert_eq!(
        run_js(
            r#"
let arr = ["banana", "apple", "cherry"];
arr.sort();
console.log(arr.join(","));
"#
        ),
        &["apple,banana,cherry"]
    );
}

#[test]
fn array_sort_numbers_correct() {
    assert_eq!(
        run_js(
            r#"
let arr = [10, 1, 21, 2];
arr.sort((a, b) => a - b);
console.log(arr.join(","));
"#
        ),
        &["1,2,10,21"]
    );
}

#[test]
fn array_sort_objects_by_property() {
    assert_eq!(
        run_js(
            r#"
let people = [
    { name: "Charlie", age: 30 },
    { name: "Alice", age: 25 },
    { name: "Bob", age: 35 }
];
people.sort((a, b) => a.age - b.age);
people.forEach(p => console.log(p.name + ":" + p.age));
"#
        ),
        &["Alice:25", "Charlie:30", "Bob:35"]
    );
}

#[test]
fn array_sort_stable() {
    assert_eq!(
        run_js(
            r#"
let items = [
    { name: "A", score: 1 },
    { name: "B", score: 2 },
    { name: "C", score: 1 },
    { name: "D", score: 2 }
];
items.sort((a, b) => a.score - b.score);
console.log(items.map(i => i.name).join(","));
"#
        ),
        &["A,C,B,D"]
    );
}

// ===================================================================
// NUMBER METHODS
// ===================================================================

#[test]
fn number_is_integer() {
    assert_eq!(
        run_js(
            r#"
console.log(Number.isInteger(42));
console.log(Number.isInteger(42.0));
console.log(Number.isInteger(42.5));
"#
        ),
        &["true", "true", "false"]
    );
}

#[test]
fn number_is_finite() {
    assert_eq!(
        run_js(
            r#"
console.log(Number.isFinite(42));
console.log(Number.isFinite(Infinity));
console.log(Number.isFinite(NaN));
"#
        ),
        &["true", "false", "false"]
    );
}

#[test]
fn number_is_nan() {
    assert_eq!(
        run_js(
            r#"
console.log(Number.isNaN(NaN));
console.log(Number.isNaN(42));
console.log(Number.isNaN("NaN"));
"#
        ),
        &["true", "false", "false"]
    );
}

#[test]
fn number_parse_int_float() {
    assert_eq!(
        run_js(
            r#"
console.log(Number.parseInt("42px"));
console.log(Number.parseFloat("3.14xyz"));
"#
        ),
        &["42", "3.14"]
    );
}

#[test]
fn number_tofixed() {
    assert_eq!(
        run_js(
            r#"
let n = 3.14159;
console.log(n.toFixed(2));
console.log(n.toFixed(0));
"#
        ),
        &["3.14", "3"]
    );
}

#[test]
fn string_repeat_zero_and_many() {
    assert_eq!(
        run_js(
            r#"
console.log("ab".repeat(0));
console.log("ab".repeat(3));
"#
        ),
        &["", "ababab"]
    );
}

#[test]
fn string_pad_start_and_end() {
    assert_eq!(
        run_js(
            r#"
console.log("5".padStart(3, "0"));
console.log("5".padEnd(3, "0"));
"#
        ),
        &["005", "500"]
    );
}

#[test]
fn string_includes_with_position() {
    assert_eq!(
        run_js(
            r#"
let s = "banana";
console.log(s.includes("an", 1));
console.log(s.includes("ba", 1));
"#
        ),
        &["true", "false"]
    );
}

#[test]
fn string_startswith_and_endswith() {
    assert_eq!(
        run_js(
            r#"
let s = "javascript";
console.log(s.startsWith("java"));
console.log(s.endsWith("script"));
console.log(s.endsWith("java"));
"#
        ),
        &["true", "true", "false"]
    );
}

#[test]
fn string_split_with_limit() {
    assert_eq!(
        run_js(
            r#"
let parts = "a,b,c,d".split(",", 2);
console.log(parts.join("|"));
"#
        ),
        &["a|b"]
    );
}

#[test]
fn string_charcodeat_and_fromcharcode_roundtrip() {
    assert_eq!(
        run_js(
            r#"
let code = "A".charCodeAt(0);
console.log(code);
console.log(String.fromCharCode(code + 1));
"#
        ),
        &["65", "B"]
    );
}

#[test]
fn string_substring_swaps_arguments() {
    assert_eq!(
        run_js(
            r#"
let s = "abcdef";
console.log(s.substring(4, 1));
console.log(s.slice(4, 1));
"#
        ),
        &["bcd", ""]
    );
}

#[test]
fn string_locale_compare_basic_ordering() {
    assert_eq!(
        run_js(
            r#"
console.log("a".localeCompare("b") < 0);
console.log("b".localeCompare("a") > 0);
console.log("a".localeCompare("a") === 0);
"#
        ),
        &["true", "true", "true"]
    );
}

#[test]
fn array_flatmap_can_remove_items() {
    assert_eq!(
        run_js(
            r#"
let arr = [1, 2, 3, 4];
let result = arr.flatMap(x => x % 2 === 0 ? [] : [x]);
console.log(result.join(","));
"#
        ),
        &["1,3"]
    );
}

#[test]
fn array_copywithin_partial_range() {
    assert_eq!(
        run_js(
            r#"
let arr = [1, 2, 3, 4, 5];
arr.copyWithin(1, 3, 5);
console.log(arr.join(","));
"#
        ),
        &["1,4,5,4,5"]
    );
}

#[test]
fn array_reduce_right_with_numbers() {
    assert_eq!(
        run_js(
            r#"
let arr = [1, 2, 3];
let result = arr.reduceRight((acc, x) => acc - x, 10);
console.log(result);
"#
        ),
        &["4"]
    );
}

#[test]
fn array_entries_iterator_exposes_index_value_pairs() {
    assert_eq!(
        run_js(
            r#"
let first = ["x", "y"].entries().next().value;
console.log(first[0]);
console.log(first[1]);
"#
        ),
        &["0", "x"]
    );
}

#[test]
fn array_keys_iterator_sequence() {
    assert_eq!(
        run_js(
            r#"
let iter = ["a", "b"].keys();
console.log(iter.next().value);
console.log(iter.next().value);
console.log(iter.next().done);
"#
        ),
        &["0", "1", "true"]
    );
}

#[test]
fn array_values_iterator_sequence() {
    assert_eq!(
        run_js(
            r#"
let iter = [10, 20].values();
console.log(iter.next().value);
console.log(iter.next().value);
console.log(iter.next().done);
"#
        ),
        &["10", "20", "true"]
    );
}

#[test]
fn array_of_single_number_differs_from_constructor() {
    assert_eq!(
        run_js(
            r#"
console.log(Array.of(3).length);
console.log(new Array(3).length);
console.log(Array.of(3)[0]);
"#
        ),
        &["1", "3", "3"]
    );
}
