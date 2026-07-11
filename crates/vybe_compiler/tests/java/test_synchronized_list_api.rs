use crate::helpers::{run_in_main, run_main};

#[test]
fn synchronized_list_add_increments_size() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> backing = new java.util.ArrayList<Integer>(); java.util.List<Integer> sync = java.util.Collections.synchronizedList(backing); sync.add(5); System.out.println(sync.size());"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn synchronized_list_get_returns_added_element() {
    let out = run_main(
        r#"java.util.ArrayList<String> backing = new java.util.ArrayList<String>(); java.util.List<String> sync = java.util.Collections.synchronizedList(backing); sync.add("item"); System.out.println(sync.get(0));"#,
    );
    assert_eq!(out, vec!["item"]);
}

#[test]
fn synchronized_list_remove_by_index() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> backing = new java.util.ArrayList<Integer>(); java.util.List<Integer> sync = java.util.Collections.synchronizedList(backing); sync.add(1); sync.add(2); sync.remove(0); System.out.println(sync.get(0));"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn synchronized_list_set_replaces_element() {
    let out = run_main(
        r#"java.util.ArrayList<String> backing = new java.util.ArrayList<String>(); java.util.List<String> sync = java.util.Collections.synchronizedList(backing); sync.add("old"); sync.set(0, "new"); System.out.println(sync.get(0));"#,
    );
    assert_eq!(out, vec!["new"]);
}

#[test]
fn synchronized_list_contains_finds_element() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> backing = new java.util.ArrayList<Integer>(); java.util.List<Integer> sync = java.util.Collections.synchronizedList(backing); sync.add(42); System.out.println(sync.contains(42));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn synchronized_list_index_of_returns_position() {
    let out = run_main(
        r#"java.util.ArrayList<String> backing = new java.util.ArrayList<String>(); java.util.List<String> sync = java.util.Collections.synchronizedList(backing); sync.add("a"); sync.add("b"); System.out.println(sync.indexOf("b"));"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn synchronized_list_clear_empties_list() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> backing = new java.util.ArrayList<Integer>(); java.util.List<Integer> sync = java.util.Collections.synchronizedList(backing); sync.add(1); sync.clear(); System.out.println(sync.isEmpty());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn synchronized_list_is_empty_on_new_wrapper() {
    let out = run_main(
        r#"java.util.ArrayList<Object> backing = new java.util.ArrayList<Object>(); java.util.List<Object> sync = java.util.Collections.synchronizedList(backing); System.out.println(sync.isEmpty());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn synchronized_list_add_all_appends_collection() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> backing = new java.util.ArrayList<Integer>(); java.util.List<Integer> sync = java.util.Collections.synchronizedList(backing); sync.addAll(java.util.Arrays.asList(1, 2, 3)); System.out.println(sync.size());"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn synchronized_list_iterator_has_next() {
    let out = run_main(
        r#"java.util.ArrayList<String> backing = new java.util.ArrayList<String>(); java.util.List<String> sync = java.util.Collections.synchronizedList(backing); sync.add("x"); System.out.println(sync.iterator().hasNext());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn synchronized_list_iterator_next_returns_element() {
    let out = run_main(
        r#"java.util.ArrayList<String> backing = new java.util.ArrayList<String>(); java.util.List<String> sync = java.util.Collections.synchronizedList(backing); sync.add("val"); System.out.println(sync.iterator().next());"#,
    );
    assert_eq!(out, vec!["val"]);
}

#[test]
fn synchronized_list_to_array_length() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> backing = new java.util.ArrayList<Integer>(); java.util.List<Integer> sync = java.util.Collections.synchronizedList(backing); sync.add(1); sync.add(2); System.out.println(sync.toArray().length);"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn synchronized_list_last_index_of_finds_last_occurrence() {
    let out = run_main(
        r#"java.util.ArrayList<String> backing = new java.util.ArrayList<String>(); java.util.List<String> sync = java.util.Collections.synchronizedList(backing); sync.add("dup"); sync.add("other"); sync.add("dup"); System.out.println(sync.lastIndexOf("dup"));"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn synchronized_list_remove_object_by_value() {
    let out = run_main(
        r#"java.util.ArrayList<String> backing = new java.util.ArrayList<String>(); java.util.List<String> sync = java.util.Collections.synchronizedList(backing); sync.add("keep"); sync.add("go"); sync.remove("go"); System.out.println(sync.size());"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn synchronized_list_add_at_index_inserts() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> backing = new java.util.ArrayList<Integer>(); java.util.List<Integer> sync = java.util.Collections.synchronizedList(backing); sync.add(2); sync.add(0, 1); System.out.println(sync.get(0));"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn synchronized_list_sublist_view_size() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> backing = new java.util.ArrayList<Integer>(); java.util.List<Integer> sync = java.util.Collections.synchronizedList(backing); sync.add(1); sync.add(2); sync.add(3); System.out.println(sync.subList(0, 2).size());"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn synchronized_list_equals_same_content() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> b1 = new java.util.ArrayList<Integer>(); java.util.List<Integer> s1 = java.util.Collections.synchronizedList(b1); s1.add(1); java.util.ArrayList<Integer> b2 = new java.util.ArrayList<Integer>(); b2.add(1); System.out.println(s1.equals(b2));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn synchronized_list_hash_code_consistent() {
    let out = run_main(
        r#"java.util.ArrayList<String> backing = new java.util.ArrayList<String>(); java.util.List<String> sync = java.util.Collections.synchronizedList(backing); sync.add("h"); int h1 = sync.hashCode(); int h2 = sync.hashCode(); System.out.println(h1 == h2);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn synchronized_list_two_threads_increment_size() {
    let out = run_in_main(
        "java.util.ArrayList<Integer> backing = new java.util.ArrayList<Integer>(); java.util.List<Integer> sync = java.util.Collections.synchronizedList(backing); Thread t1 = new Thread(() -> { for (int i = 0; i < 25; i++) sync.add(i); }); Thread t2 = new Thread(() -> { for (int i = 0; i < 25; i++) sync.add(i); }); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(sync.size());",
        "",
    );
    assert_eq!(out, vec!["50"]);
}

#[test]
fn synchronized_list_retains_all_filters() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> backing = new java.util.ArrayList<Integer>(); java.util.List<Integer> sync = java.util.Collections.synchronizedList(backing); sync.add(1); sync.add(2); sync.add(3); sync.retainAll(java.util.Arrays.asList(2, 3)); System.out.println(sync.size());"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn synchronized_list_remove_all_batch() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> backing = new java.util.ArrayList<Integer>(); java.util.List<Integer> sync = java.util.Collections.synchronizedList(backing); sync.add(1); sync.add(2); sync.add(3); sync.removeAll(java.util.Arrays.asList(1, 3)); System.out.println(sync.get(0));"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn synchronized_list_contains_all_true() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> backing = new java.util.ArrayList<Integer>(); java.util.List<Integer> sync = java.util.Collections.synchronizedList(backing); sync.add(1); sync.add(2); System.out.println(sync.containsAll(java.util.Arrays.asList(1, 2)));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn synchronized_list_list_iterator_has_next() {
    let out = run_main(
        r#"java.util.ArrayList<String> backing = new java.util.ArrayList<String>(); java.util.List<String> sync = java.util.Collections.synchronizedList(backing); sync.add("a"); System.out.println(sync.listIterator().hasNext());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn synchronized_list_list_iterator_next_index() {
    let out = run_main(
        r#"java.util.ArrayList<String> backing = new java.util.ArrayList<String>(); java.util.List<String> sync = java.util.Collections.synchronizedList(backing); sync.add("a"); java.util.ListIterator<String> it = sync.listIterator(); it.next(); System.out.println(it.nextIndex());"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn synchronized_list_add_null_element() {
    let out = run_main(
        r#"java.util.ArrayList<String> backing = new java.util.ArrayList<String>(); java.util.List<String> sync = java.util.Collections.synchronizedList(backing); sync.add(null); System.out.println(sync.get(0) == null);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn synchronized_list_get_first_after_multiple_adds() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> backing = new java.util.ArrayList<Integer>(); java.util.List<Integer> sync = java.util.Collections.synchronizedList(backing); sync.add(10); sync.add(20); sync.add(30); System.out.println(sync.get(0)); System.out.println(sync.get(2));"#,
    );
    assert_eq!(out, vec!["10", "30"]);
}

#[test]
fn synchronized_list_size_after_remove_all() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> backing = new java.util.ArrayList<Integer>(); java.util.List<Integer> sync = java.util.Collections.synchronizedList(backing); sync.add(1); sync.add(2); sync.removeAll(java.util.Arrays.asList(1, 2)); System.out.println(sync.size());"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn synchronized_list_wrapper_same_size_as_backing() {
    let out = run_main(
        r#"java.util.ArrayList<String> backing = new java.util.ArrayList<String>(); java.util.List<String> sync = java.util.Collections.synchronizedList(backing); sync.add("x"); sync.add("y"); System.out.println(sync.size()); System.out.println(backing.size());"#,
    );
    assert_eq!(out, vec!["2", "2"]);
}

#[test]
fn synchronized_list_for_each_loop_prints_elements() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> backing = new java.util.ArrayList<Integer>(); java.util.List<Integer> sync = java.util.Collections.synchronizedList(backing); sync.add(3); sync.add(7); int sum = 0; for (int v : sync) { sum += v; } System.out.println(sum);"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn synchronized_list_add_returns_true() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> backing = new java.util.ArrayList<Integer>(); java.util.List<Integer> sync = java.util.Collections.synchronizedList(backing); System.out.println(sync.add(99));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn synchronized_list_remove_returns_removed_element() {
    let out = run_main(
        r#"java.util.ArrayList<String> backing = new java.util.ArrayList<String>(); java.util.List<String> sync = java.util.Collections.synchronizedList(backing); sync.add("gone"); System.out.println(sync.remove(0));"#,
    );
    assert_eq!(out, vec!["gone"]);
}

#[test]
fn synchronized_list_set_returns_previous_value() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> backing = new java.util.ArrayList<Integer>(); java.util.List<Integer> sync = java.util.Collections.synchronizedList(backing); sync.add(1); System.out.println(sync.set(0, 2));"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn synchronized_list_to_string_nonempty() {
    let out = run_main(
        r#"java.util.ArrayList<String> backing = new java.util.ArrayList<String>(); java.util.List<String> sync = java.util.Collections.synchronizedList(backing); sync.add("a"); System.out.println(sync.toString().length() > 0);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn synchronized_list_stream_count() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> backing = new java.util.ArrayList<Integer>(); java.util.List<Integer> sync = java.util.Collections.synchronizedList(backing); sync.add(1); sync.add(2); System.out.println(sync.stream().count());"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn synchronized_list_spliterator_estimate_size() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> backing = new java.util.ArrayList<Integer>(); java.util.List<Integer> sync = java.util.Collections.synchronizedList(backing); sync.add(1); sync.add(2); sync.add(3); System.out.println(sync.spliterator().estimateSize());"#,
    );
    assert_eq!(out, vec!["3"]);
}
