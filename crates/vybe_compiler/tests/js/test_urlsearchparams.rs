crate::js_cases! {
    urlsearchparams_reads_existing_values => {
        r#"
const params = new URLSearchParams("q=vybe&page=2");
console.log(params.get("q"));
console.log(params.get("page"));
"#,
        ["vybe", "2"]
    };

    urlsearchparams_append_preserves_multiple_values => {
        r#"
const params = new URLSearchParams();
params.append("tag", "js");
params.append("tag", "tests");
console.log(params.getAll("tag").join(","));
"#,
        ["js,tests"]
    };

    urlsearchparams_set_replaces_existing_values => {
        r#"
const params = new URLSearchParams("mode=old&mode=stale");
params.set("mode", "fresh");
console.log(params.getAll("mode").join(","));
"#,
        ["fresh"]
    };

    urlsearchparams_delete_removes_key => {
        r#"
const params = new URLSearchParams("a=1&b=2&a=3");
params.delete("a");
console.log(params.has("a"));
console.log(params.get("b"));
"#,
        ["false", "2"]
    };

    urlsearchparams_sort_orders_pairs_by_key => {
        r#"
const params = new URLSearchParams("z=1&a=2&m=3");
params.sort();
console.log(params.toString());
"#,
        ["a=2&m=3&z=1"]
    };

    urlsearchparams_tostring_percent_encodes_spaces => {
        r#"
const params = new URLSearchParams();
params.set("query", "two words");
console.log(params.toString());
"#,
        ["query=two+words"]
    };

    urlsearchparams_iterates_entries_in_insertion_order => {
        r#"
const params = new URLSearchParams("a=1&b=2&a=3");
console.log([...params.entries()].map(([k, v]) => k + ":" + v).join(","));
"#,
        ["a:1,b:2,a:3"]
    };

    urlsearchparams_constructs_from_object_record => {
        r#"
const params = new URLSearchParams({ lang: "js", level: "advanced" });
console.log(params.get("lang"));
console.log(params.get("level"));
"#,
        ["js", "advanced"]
    };

    urlsearchparams_constructs_from_sequence_pairs => {
        r#"
const params = new URLSearchParams([["x", "1"], ["y", "2"]]);
console.log(params.toString());
"#,
        ["x=1&y=2"]
    };
}