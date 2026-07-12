crate::js_cases! {
    urlsearchparams_empty_string_has_no_entries => {
        r#"
const p = new URLSearchParams("");
console.log([...p].length);
"#,
        ["0"]
    };

    urlsearchparams_leading_question_mark_is_ignored => {
        r#"
const p = new URLSearchParams("?a=1&b=2");
console.log(p.get("a"));
console.log(p.get("b"));
"#,
        ["1", "2"]
    };

    urlsearchparams_get_missing_returns_null => {
        r#"
const p = new URLSearchParams("a=1");
console.log(p.get("x") === null);
"#,
        ["true"]
    };

    urlsearchparams_getall_missing_returns_empty_array => {
        r#"
const p = new URLSearchParams("a=1");
console.log(p.getAll("x").length);
"#,
        ["0"]
    };

    urlsearchparams_set_absent_key_adds_new_pair => {
        r#"
const p = new URLSearchParams("a=1");
p.set("b", "2");
console.log(p.toString());
"#,
        ["a=1&b=2"]
    };

    urlsearchparams_append_preserves_existing_order => {
        r#"
const p = new URLSearchParams("a=1&b=2");
p.append("a", "3");
console.log([...p.entries()].map(([k,v]) => k + ":" + v).join(","));
"#,
        ["a:1,b:2,a:3"]
    };

    urlsearchparams_delete_missing_key_is_noop => {
        r#"
const p = new URLSearchParams("a=1");
p.delete("x");
console.log(p.toString());
"#,
        ["a=1"]
    };

    urlsearchparams_keys_iterator_yields_names => {
        r#"
const p = new URLSearchParams("a=1&b=2&a=3");
console.log([...p.keys()].join(","));
"#,
        ["a,b,a"]
    };

    urlsearchparams_values_iterator_yields_values => {
        r#"
const p = new URLSearchParams("a=1&b=2&a=3");
console.log([...p.values()].join(","));
"#,
        ["1,2,3"]
    };

    urlsearchparams_entries_iterator_yields_pairs => {
        r#"
const p = new URLSearchParams("a=1&b=2");
console.log([...p.entries()].map(([k,v]) => k + ":" + v).join(","));
"#,
        ["a:1,b:2"]
    };

    urlsearchparams_foreach_receives_name_value_pairs => {
        r#"
const p = new URLSearchParams("a=1&b=2");
const out = [];
p.forEach((value, key) => out.push(key + ":" + value));
console.log(out.join(","));
"#,
        ["a:1,b:2"]
    };

    urlsearchparams_foreach_honors_this_arg => {
        r#"
const p = new URLSearchParams("a=1&b=2");
const ctx = { prefix: ">" };
const out = [];
p.forEach(function(value, key) { out.push(this.prefix + key + value); }, ctx);
console.log(out.join(","));
"#,
        [">a1,>b2"]
    };

    urlsearchparams_plus_decodes_to_space => {
        r#"
const p = new URLSearchParams("q=hello+world");
console.log(p.get("q"));
"#,
        ["hello world"]
    };

    urlsearchparams_percent_encoding_decodes_utf8 => {
        r#"
const p = new URLSearchParams("q=%C3%A9");
console.log(p.get("q"));
"#,
        ["é"]
    };

    urlsearchparams_empty_value_is_preserved => {
        r#"
const p = new URLSearchParams("a=");
console.log(p.get("a") === "");
"#,
        ["true"]
    };

    urlsearchparams_empty_key_is_preserved => {
        r#"
const p = new URLSearchParams("=value");
console.log(p.get(""));
"#,
        ["value"]
    };

    urlsearchparams_duplicate_empty_keys_preserve_all_values => {
        r#"
const p = new URLSearchParams("=a&=b");
console.log(p.getAll("").join(","));
"#,
        ["a,b"]
    };

    urlsearchparams_tostring_on_empty_params_is_empty => {
        r#"
console.log(new URLSearchParams().toString() === "");
"#,
        ["true"]
    };

    urlsearchparams_numeric_values_are_stringified => {
        r#"
const p = new URLSearchParams();
p.append("n", 42);
console.log(p.get("n"));
"#,
        ["42"]
    };

    urlsearchparams_sort_is_stable_for_equal_keys => {
        r#"
const p = new URLSearchParams("b=1&a=2&a=3");
p.sort();
console.log(p.toString());
"#,
        ["a=2&a=3&b=1"]
    };

    urlsearchparams_default_iterator_matches_entries => {
        r#"
const p = new URLSearchParams("a=1&b=2");
console.log([...p].map(([k,v]) => k + ":" + v).join(","));
"#,
        ["a:1,b:2"]
    };

    urlsearchparams_has_true_for_present_key => {
        r#"
const p = new URLSearchParams("a=1");
console.log(p.has("a"));
"#,
        ["true"]
    };

    urlsearchparams_has_false_for_missing_key => {
        r#"
const p = new URLSearchParams("a=1");
console.log(p.has("b"));
"#,
        ["false"]
    };

    urlsearchparams_constructor_from_pairs_preserves_duplicate_names => {
        r#"
const p = new URLSearchParams([["a", "1"], ["a", "2"]]);
console.log(p.getAll("a").join(","));
"#,
        ["1,2"]
    };
}
