/// java.util.concurrent.CopyOnWriteArrayList — snapshot iterator semantics.
use crate::helpers::{run_in_main, run_main};

#[test]
fn copy_on_write_list_add_and_get_round_trip() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); list.add(10); System.out.println(list.get(0));",
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn copy_on_write_list_new_list_is_empty() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<String> list = new java.util.concurrent.CopyOnWriteArrayList<String>(); System.out.println(list.isEmpty());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn copy_on_write_list_size_after_three_adds() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); list.add(1); list.add(2); list.add(3); System.out.println(list.size());",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn copy_on_write_list_add_at_index_inserts_in_middle() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); list.add(1); list.add(3); list.add(1, 2); System.out.println(list.get(1));",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn copy_on_write_list_set_replaces_element_at_index() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); list.add(1); list.add(2); list.set(0, 9); System.out.println(list.get(0));",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn copy_on_write_list_set_returns_previous_value() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); list.add(5); System.out.println(list.set(0, 8));",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn copy_on_write_list_remove_by_index() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); list.add(1); list.add(2); list.remove(0); System.out.println(list.get(0));",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn copy_on_write_list_remove_by_object() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); list.add(7); list.add(8); System.out.println(list.remove(Integer.valueOf(7))); System.out.println(list.size());",
    );
    assert_eq!(out, vec!["true", "1"]);
}

#[test]
fn copy_on_write_list_contains_finds_added_element() {
    let out = run_main(
        r#"java.util.concurrent.CopyOnWriteArrayList<String> list = new java.util.concurrent.CopyOnWriteArrayList<String>(); list.add("vybe"); System.out.println(list.contains("vybe"));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn copy_on_write_list_index_of_returns_first_position() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); list.add(3); list.add(4); list.add(3); System.out.println(list.indexOf(3));",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn copy_on_write_list_last_index_of_returns_final_position() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); list.add(3); list.add(4); list.add(3); System.out.println(list.lastIndexOf(3));",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn copy_on_write_list_add_all_appends_collection() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); java.util.ArrayList<Integer> other = new java.util.ArrayList<Integer>(); other.add(1); other.add(2); list.addAll(other); System.out.println(list.size()); System.out.println(list.get(1));",
    );
    assert_eq!(out, vec!["2", "2"]);
}

#[test]
fn copy_on_write_list_add_all_at_index_inserts_block() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); list.add(1); list.add(4); java.util.ArrayList<Integer> mid = new java.util.ArrayList<Integer>(); mid.add(2); mid.add(3); list.addAll(1, mid); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn copy_on_write_list_add_if_absent_inserts_new_element() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); System.out.println(list.addIfAbsent(5)); System.out.println(list.contains(5));",
    );
    assert_eq!(out, vec!["true", "true"]);
}

#[test]
fn copy_on_write_list_add_if_absent_skips_duplicate() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); list.add(5); System.out.println(list.addIfAbsent(5)); System.out.println(list.size());",
    );
    assert_eq!(out, vec!["false", "1"]);
}

#[test]
fn copy_on_write_list_iterator_traverses_in_order() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); list.add(1); list.add(2); list.add(3); java.util.Iterator<Integer> it = list.iterator(); System.out.println(it.next()); System.out.println(it.next()); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn copy_on_write_list_iterator_does_not_support_remove() {
    let out = run_in_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); list.add(1); java.util.Iterator<Integer> it = list.iterator(); it.next(); try { it.remove(); System.out.println(\"ok\"); } catch (UnsupportedOperationException e) { System.out.println(\"unsupported\"); }",
        "",
    );
    assert_eq!(out, vec!["unsupported"]);
}

#[test]
fn copy_on_write_list_list_iterator_supports_add() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); list.add(1); list.add(3); java.util.ListIterator<Integer> it = list.listIterator(); it.next(); it.add(2); System.out.println(list.get(1));",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn copy_on_write_list_to_array_has_correct_length() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); list.add(1); list.add(2); Object[] arr = list.toArray(); System.out.println(arr.length);",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn copy_on_write_list_to_array_typed_copies_elements() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<String> list = new java.util.concurrent.CopyOnWriteArrayList<String>(); list.add(\"a\"); list.add(\"b\"); String[] arr = list.toArray(new String[0]); System.out.println(arr[1]);",
    );
    assert_eq!(out, vec!["b"]);
}

#[test]
fn copy_on_write_list_clear_removes_all_elements() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); list.add(1); list.add(2); list.clear(); System.out.println(list.isEmpty());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn copy_on_write_list_remove_all_eliminates_matching() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); list.add(1); list.add(2); list.add(3); java.util.ArrayList<Integer> drop = new java.util.ArrayList<Integer>(); drop.add(2); System.out.println(list.removeAll(drop)); System.out.println(list.size());",
    );
    assert_eq!(out, vec!["true", "2"]);
}

#[test]
fn copy_on_write_list_retain_all_keeps_only_matches() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); list.add(1); list.add(2); list.add(3); java.util.ArrayList<Integer> keep = new java.util.ArrayList<Integer>(); keep.add(2); keep.add(3); list.retainAll(keep); System.out.println(list.size()); System.out.println(list.get(0));",
    );
    assert_eq!(out, vec!["2", "2"]);
}

#[test]
fn copy_on_write_list_contains_all_checks_superset() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); list.add(1); list.add(2); list.add(3); java.util.ArrayList<Integer> probe = new java.util.ArrayList<Integer>(); probe.add(2); probe.add(3); System.out.println(list.containsAll(probe));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn copy_on_write_list_sub_list_is_view_of_range() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); list.add(1); list.add(2); list.add(3); java.util.List<Integer> sub = list.subList(1, 3); System.out.println(sub.size()); System.out.println(sub.get(0));",
    );
    assert_eq!(out, vec!["2", "2"]);
}

#[test]
fn copy_on_write_list_constructor_from_collection() {
    let out = run_main(
        "java.util.ArrayList<Integer> src = new java.util.ArrayList<Integer>(); src.add(4); src.add(5); java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(src); System.out.println(list.get(0)); System.out.println(list.size());",
    );
    assert_eq!(out, vec!["4", "2"]);
}

#[test]
fn copy_on_write_list_add_all_no_return_false_when_unchanged() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); java.util.ArrayList<Integer> empty = new java.util.ArrayList<Integer>(); System.out.println(list.addAll(empty));",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn copy_on_write_list_get_first_and_last_via_indices() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); list.add(10); list.add(20); System.out.println(list.get(0)); System.out.println(list.get(list.size() - 1));",
    );
    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn copy_on_write_list_iterator_has_next_false_on_empty() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); System.out.println(list.iterator().hasNext());",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn copy_on_write_list_list_iterator_previous_index() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); list.add(1); java.util.ListIterator<Integer> it = list.listIterator(); it.next(); System.out.println(it.previousIndex());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn copy_on_write_list_remove_returns_removed_element() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); list.add(9); System.out.println(list.remove(0));",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn copy_on_write_list_add_at_end_via_index() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); list.add(0, 7); System.out.println(list.get(0));",
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn copy_on_write_list_concurrent_read_while_write_visible_after() {
    let types = r#"
        static java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>();
        static int snapshotSize;
    "#;
    let out = run_in_main(
        "list.add(1); Thread writer = new Thread(() -> list.add(2)); Thread reader = new Thread(() -> snapshotSize = list.size()); writer.start(); reader.start(); writer.join(); reader.join(); System.out.println(list.size());",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn copy_on_write_list_thread_safe_add_from_two_threads() {
    let types = r#"
        static java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>();
    "#;
    let out = run_in_main(
        "Thread t1 = new Thread(() -> list.add(1)); Thread t2 = new Thread(() -> list.add(2)); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(list.size());",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn copy_on_write_list_equals_same_content_different_instance() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> a = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); a.add(1); java.util.concurrent.CopyOnWriteArrayList<Integer> b = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); b.add(1); System.out.println(a.equals(b));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn copy_on_write_list_hash_code_consistent_for_equal_lists() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> a = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); a.add(1); java.util.concurrent.CopyOnWriteArrayList<Integer> b = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); b.add(1); System.out.println(a.hashCode() == b.hashCode());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn copy_on_write_list_remove_object_returns_false_when_absent() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); list.add(1); System.out.println(list.remove(Integer.valueOf(9)));",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn copy_on_write_list_index_of_absent_returns_negative_one() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); list.add(1); System.out.println(list.indexOf(99));",
    );
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn copy_on_write_list_list_iterator_next_index_after_first() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); list.add(1); list.add(2); java.util.ListIterator<Integer> it = list.listIterator(); it.next(); System.out.println(it.nextIndex());",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn copy_on_write_list_add_all_return_true_when_modified() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); java.util.ArrayList<Integer> batch = new java.util.ArrayList<Integer>(); batch.add(1); System.out.println(list.addAll(batch));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn copy_on_write_list_for_each_prints_each_element() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); list.add(2); list.add(4); list.forEach(n -> System.out.println(n));",
    );
    assert_eq!(out, vec!["2", "4"]);
}

#[test]
fn copy_on_write_list_stream_count_matches_size() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); list.add(1); list.add(2); list.add(3); System.out.println(list.stream().count());",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn copy_on_write_list_remove_all_on_empty_returns_false() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); java.util.ArrayList<Integer> drop = new java.util.ArrayList<Integer>(); drop.add(1); System.out.println(list.removeAll(drop));",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn copy_on_write_list_retain_all_on_empty_is_noop() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); java.util.ArrayList<Integer> keep = new java.util.ArrayList<Integer>(); keep.add(1); System.out.println(list.retainAll(keep)); System.out.println(list.isEmpty());",
    );
    assert_eq!(out, vec!["false", "true"]);
}

#[test]
fn copy_on_write_list_add_if_absent_thread_safe_only_one_insert() {
    let types = r#"
        static java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>();
        static int inserts = 0;
    "#;
    let out = run_in_main(
        "Thread t1 = new Thread(() -> { if (list.addIfAbsent(7)) inserts++; }); Thread t2 = new Thread(() -> { if (list.addIfAbsent(7)) inserts++; }); t1.start(); t2.start(); t1.join(); t2.join(); System.out.println(inserts); System.out.println(list.size());",
        types,
    );
    assert_eq!(out, vec!["1", "1"]);
}

#[test]
fn copy_on_write_list_list_iterator_at_index_starts_mid_list() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); list.add(1); list.add(2); list.add(3); java.util.ListIterator<Integer> it = list.listIterator(1); System.out.println(it.next());",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn copy_on_write_list_remove_all_clears_everything_when_all_match() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); list.add(5); list.add(5); java.util.ArrayList<Integer> all = new java.util.ArrayList<Integer>(); all.add(5); list.removeAll(all); System.out.println(list.isEmpty());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn copy_on_write_list_contains_all_false_when_missing_element() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); list.add(1); java.util.ArrayList<Integer> probe = new java.util.ArrayList<Integer>(); probe.add(1); probe.add(2); System.out.println(list.containsAll(probe));",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn copy_on_write_list_set_at_last_index() {
    let out = run_main(
        "java.util.concurrent.CopyOnWriteArrayList<Integer> list = new java.util.concurrent.CopyOnWriteArrayList<Integer>(); list.add(1); list.add(2); list.set(1, 9); System.out.println(list.get(1));",
    );
    assert_eq!(out, vec!["9"]);
}
