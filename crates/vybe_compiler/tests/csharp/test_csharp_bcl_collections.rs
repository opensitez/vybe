//! BCL collection types — each test covers the defining contract of one container.
use super::helpers::run_csharp;

#[test]
fn stack_pop_returns_most_recently_pushed_element() {
    assert_eq!(
        run_csharp(
            r#"
var stack = new System.Collections.Generic.Stack<int>();
stack.Push(1);
stack.Push(2);
Console.WriteLine(stack.Pop());
"#
        ),
        &["2"]
    );
}

#[test]
fn queue_dequeue_returns_oldest_enqueued_element() {
    assert_eq!(
        run_csharp(
            r#"
var queue = new System.Collections.Generic.Queue<int>();
queue.Enqueue(1);
queue.Enqueue(2);
Console.WriteLine(queue.Dequeue());
"#
        ),
        &["1"]
    );
}

#[test]
fn hashset_add_returns_false_for_duplicate_element() {
    assert_eq!(
        run_csharp(
            r#"
var set = new System.Collections.Generic.HashSet<int>();
Console.WriteLine(set.Add(1));
Console.WriteLine(set.Add(1));
"#
        ),
        &["True", "False"]
    );
}

#[test]
fn hashset_union_with_merges_distinct_elements_from_other_set() {
    assert_eq!(
        run_csharp(
            r#"
var left = new System.Collections.Generic.HashSet<int> { 1, 2 };
var right = new System.Collections.Generic.HashSet<int> { 2, 3 };
left.UnionWith(right);
Console.WriteLine(left.Count);
"#
        ),
        &["3"]
    );
}

#[test]
fn linked_list_add_first_inserts_before_current_head() {
    assert_eq!(
        run_csharp(
            r#"
var list = new System.Collections.Generic.LinkedList<int>();
list.AddLast(2);
list.AddFirst(1);
Console.WriteLine(list.First.Value);
"#
        ),
        &["1"]
    );
}

#[test]
fn sorted_dictionary_enumerator_yields_keys_in_ascending_order() {
    assert_eq!(
        run_csharp(
            r#"
var map = new System.Collections.Generic.SortedDictionary<int, string>();
map[3] = "c";
map[1] = "a";
int firstKey = 0;
foreach (var pair in map) { firstKey = pair.Key; break; }
Console.WriteLine(firstKey);
"#
        ),
        &["1"]
    );
}

#[test]
fn read_only_list_view_reflects_backing_list_mutations() {
    assert_eq!(
        run_csharp(
            r#"
var backing = new System.Collections.Generic.List<int> { 1 };
var view = backing.AsReadOnly();
backing.Add(2);
Console.WriteLine(view.Count);
"#
        ),
        &["2"]
    );
}

#[test]
fn dictionary_with_string_comparer_uses_case_insensitive_key_lookup() {
    assert_eq!(
        run_csharp(
            r#"
var map = new System.Collections.Generic.Dictionary<string, int>(
    System.StringComparer.OrdinalIgnoreCase);
map["Key"] = 7;
Console.WriteLine(map.ContainsKey("key"));
"#
        ),
        &["True"]
    );
}
