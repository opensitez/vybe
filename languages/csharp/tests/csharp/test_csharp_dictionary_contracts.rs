//! `Dictionary<K,V>` lookup, update, and duplicate-key behavior.
use super::helpers::run_csharp;

#[test]
fn indexer_read_returns_stored_value_for_existing_key() {
    assert_eq!(
        run_csharp(
            r#"
using System.Collections.Generic;
var map = new Dictionary<string, int> { ["pi"] = 3 };
Console.WriteLine(map["pi"]);
"#
        ),
        &["3"]
    );
}

#[test]
fn indexer_assignment_updates_existing_entry_without_growing_count() {
    assert_eq!(
        run_csharp(
            r#"
using System.Collections.Generic;
var map = new Dictionary<string, int> { ["x"] = 1 };
map["x"] = 9;
Console.WriteLine(map["x"]);
Console.WriteLine(map.Count);
"#
        ),
        &["9", "1"]
    );
}

#[test]
fn try_get_value_reports_absence_without_adding_entries() {
    assert_eq!(
        run_csharp(
            r#"
using System.Collections.Generic;
var map = new Dictionary<string, int> { ["a"] = 1 };
bool found = map.TryGetValue("missing", out var value);
Console.WriteLine(found ? "Y" : "N");
Console.WriteLine(map.Count);
"#
        ),
        &["N", "1"]
    );
}

#[test]
fn contains_key_reflects_add_and_remove_lifecycle() {
    assert_eq!(
        run_csharp(
            r#"
using System.Collections.Generic;
var map = new Dictionary<int, string>();
map[1] = "one";
Console.WriteLine(map.ContainsKey(1) ? "Y" : "N");
map.Remove(1);
Console.WriteLine(map.ContainsKey(1) ? "Y" : "N");
"#
        ),
        &["Y", "N"]
    );
}

#[test]
fn foreach_over_dictionary_emits_key_value_pairs_in_insertion_order() {
    assert_eq!(
        run_csharp(
            r#"
using System.Collections.Generic;
var map = new Dictionary<string, int> {
    ["b"] = 2,
    ["a"] = 1,
    ["c"] = 3
};
foreach (var entry in map) {
    Console.WriteLine(entry.Key + ":" + entry.Value);
}
"#
        ),
        &["b:2", "a:1", "c:3"]
    );
}
