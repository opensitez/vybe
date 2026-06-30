use crate::helpers::run_main;

#[test]
fn vector_default_constructor_empty_size() {
    let out = run_main(r#"java.util.Vector<String> v = new java.util.Vector<String>(); System.out.println(v.size());"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn vector_capacity_constructor_sets_capacity() {
    let out = run_main(r#"java.util.Vector<Integer> v = new java.util.Vector<Integer>(10); System.out.println(v.capacity());"#);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn vector_add_appends_element() {
    let out = run_main(r#"java.util.Vector<String> v = new java.util.Vector<String>(); v.add("alpha"); System.out.println(v.get(0));"#);
    assert_eq!(out, vec!["alpha"]);
}

#[test]
fn vector_add_returns_true() {
    let out = run_main(r#"java.util.Vector<Integer> v = new java.util.Vector<Integer>(); System.out.println(v.add(7));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn vector_insert_at_index_shifts_elements() {
    let out = run_main(r#"java.util.Vector<String> v = new java.util.Vector<String>(); v.add("b"); v.add(0, "a"); System.out.println(v.get(0)); System.out.println(v.get(1));"#);
    assert_eq!(out, vec!["a", "b"]);
}

#[test]
fn vector_element_at_returns_value() {
    let out = run_main(r#"java.util.Vector<Integer> v = new java.util.Vector<Integer>(); v.add(42); System.out.println(v.elementAt(0));"#);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn vector_first_element_returns_head() {
    let out = run_main(r#"java.util.Vector<String> v = new java.util.Vector<String>(); v.add("first"); v.add("second"); System.out.println(v.firstElement());"#);
    assert_eq!(out, vec!["first"]);
}

#[test]
fn vector_last_element_returns_tail() {
    let out = run_main(r#"java.util.Vector<String> v = new java.util.Vector<String>(); v.add("first"); v.add("last"); System.out.println(v.lastElement());"#);
    assert_eq!(out, vec!["last"]);
}

#[test]
fn vector_remove_by_index() {
    let out = run_main(r#"java.util.Vector<Integer> v = new java.util.Vector<Integer>(); v.add(1); v.add(2); v.remove(0); System.out.println(v.get(0)); System.out.println(v.size());"#);
    assert_eq!(out, vec!["2", "1"]);
}

#[test]
fn vector_remove_by_object() {
    let out = run_main(r#"java.util.Vector<String> v = new java.util.Vector<String>(); v.add("keep"); v.add("drop"); v.remove("drop"); System.out.println(v.size());"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn vector_set_replaces_at_index() {
    let out = run_main(r#"java.util.Vector<Integer> v = new java.util.Vector<Integer>(); v.add(1); v.set(0, 9); System.out.println(v.get(0));"#);
    assert_eq!(out, vec!["9"]);
}

#[test]
fn vector_set_size_truncates() {
    let out = run_main(r#"java.util.Vector<Integer> v = new java.util.Vector<Integer>(); v.add(1); v.add(2); v.add(3); v.setSize(1); System.out.println(v.size());"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn vector_set_size_grows_with_nulls() {
    let out = run_main(r#"java.util.Vector<String> v = new java.util.Vector<String>(); v.setSize(3); System.out.println(v.size());"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn vector_capacity_grows_after_adds() {
    let out = run_main(r#"java.util.Vector<Integer> v = new java.util.Vector<Integer>(); for (int i = 0; i < 20; i++) v.add(i); System.out.println(v.capacity() >= 20);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn vector_trim_to_size_reduces_capacity() {
    let out = run_main(r#"java.util.Vector<Integer> v = new java.util.Vector<Integer>(20); v.add(1); v.trimToSize(); System.out.println(v.capacity());"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn vector_ensure_capacity_increases() {
    let out = run_main(r#"java.util.Vector<Integer> v = new java.util.Vector<Integer>(); v.ensureCapacity(50); System.out.println(v.capacity() >= 50);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn vector_contains_finds_element() {
    let out = run_main(r#"java.util.Vector<String> v = new java.util.Vector<String>(); v.add("findme"); System.out.println(v.contains("findme"));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn vector_index_of_returns_position() {
    let out = run_main(r#"java.util.Vector<String> v = new java.util.Vector<String>(); v.add("a"); v.add("b"); System.out.println(v.indexOf("b"));"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn vector_clear_empties_collection() {
    let out = run_main(r#"java.util.Vector<Integer> v = new java.util.Vector<Integer>(); v.add(1); v.clear(); System.out.println(v.isEmpty());"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn vector_is_empty_true_when_new() {
    let out = run_main(r#"java.util.Vector<Object> v = new java.util.Vector<Object>(); System.out.println(v.isEmpty());"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn vector_clone_copies_elements() {
    let out = run_main(r#"java.util.Vector<Integer> v = new java.util.Vector<Integer>(); v.add(5); java.util.Vector<Integer> c = (java.util.Vector<Integer>) v.clone(); System.out.println(c.get(0));"#);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn vector_elements_enumeration_has_next() {
    let out = run_main(r#"java.util.Vector<String> v = new java.util.Vector<String>(); v.add("x"); java.util.Enumeration<String> e = v.elements(); System.out.println(e.hasMoreElements());"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn vector_elements_enumeration_next_element() {
    let out = run_main(r#"java.util.Vector<String> v = new java.util.Vector<String>(); v.add("val"); java.util.Enumeration<String> e = v.elements(); System.out.println(e.nextElement());"#);
    assert_eq!(out, vec!["val"]);
}

#[test]
fn vector_add_all_collection() {
    let out = run_main(r#"java.util.Vector<Integer> v = new java.util.Vector<Integer>(); java.util.List<Integer> other = java.util.Arrays.asList(1, 2, 3); v.addAll(other); System.out.println(v.size());"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn vector_sublist_view_size() {
    let out = run_main(r#"java.util.Vector<Integer> v = new java.util.Vector<Integer>(); v.add(1); v.add(2); v.add(3); java.util.List<Integer> sub = v.subList(1, 3); System.out.println(sub.size());"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn vector_to_array_length_matches_size() {
    let out = run_main(r#"java.util.Vector<String> v = new java.util.Vector<String>(); v.add("a"); v.add("b"); Object[] arr = v.toArray(); System.out.println(arr.length);"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn hashtable_default_constructor_empty() {
    let out = run_main(r#"java.util.Hashtable<String, Integer> h = new java.util.Hashtable<String, Integer>(); System.out.println(h.isEmpty());"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn hashtable_put_and_get_roundtrip() {
    let out = run_main(r#"java.util.Hashtable<String, String> h = new java.util.Hashtable<String, String>(); h.put("key", "value"); System.out.println(h.get("key"));"#);
    assert_eq!(out, vec!["value"]);
}

#[test]
fn hashtable_put_returns_previous_value() {
    let out = run_main(r#"java.util.Hashtable<String, Integer> h = new java.util.Hashtable<String, Integer>(); h.put("k", 1); System.out.println(h.put("k", 2));"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn hashtable_put_null_key_rejected() {
    let out = run_main(r#"try { java.util.Hashtable<String, String> h = new java.util.Hashtable<String, String>(); h.put(null, "v"); System.out.println("fail"); } catch (NullPointerException e) { System.out.println("npe"); }"#);
    assert_eq!(out, vec!["npe"]);
}

#[test]
fn hashtable_contains_key_true() {
    let out = run_main(r#"java.util.Hashtable<String, Integer> h = new java.util.Hashtable<String, Integer>(); h.put("a", 1); System.out.println(h.containsKey("a"));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn hashtable_contains_value_true() {
    let out = run_main(r#"java.util.Hashtable<String, Integer> h = new java.util.Hashtable<String, Integer>(); h.put("a", 99); System.out.println(h.containsValue(99));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn hashtable_remove_returns_removed_value() {
    let out = run_main(r#"java.util.Hashtable<String, String> h = new java.util.Hashtable<String, String>(); h.put("x", "y"); System.out.println(h.remove("x"));"#);
    assert_eq!(out, vec!["y"]);
}

#[test]
fn hashtable_clear_empties_map() {
    let out = run_main(r#"java.util.Hashtable<String, Integer> h = new java.util.Hashtable<String, Integer>(); h.put("a", 1); h.clear(); System.out.println(h.size());"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn hashtable_size_counts_entries() {
    let out = run_main(r#"java.util.Hashtable<String, Integer> h = new java.util.Hashtable<String, Integer>(); h.put("a", 1); h.put("b", 2); System.out.println(h.size());"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn hashtable_keys_enumeration() {
    let out = run_main(r#"java.util.Hashtable<String, Integer> h = new java.util.Hashtable<String, Integer>(); h.put("k", 1); java.util.Enumeration<String> keys = h.keys(); System.out.println(keys.nextElement());"#);
    assert_eq!(out, vec!["k"]);
}

#[test]
fn hashtable_elements_enumeration_values() {
    let out = run_main(r#"java.util.Hashtable<String, Integer> h = new java.util.Hashtable<String, Integer>(); h.put("k", 42); java.util.Enumeration<Integer> vals = h.elements(); System.out.println(vals.nextElement());"#);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn hashtable_entry_set_size() {
    let out = run_main(r#"java.util.Hashtable<String, Integer> h = new java.util.Hashtable<String, Integer>(); h.put("a", 1); h.put("b", 2); System.out.println(h.entrySet().size());"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn hashtable_put_all_merges_maps() {
    let out = run_main(r#"java.util.Hashtable<String, Integer> h = new java.util.Hashtable<String, Integer>(); java.util.Map<String, Integer> m = new java.util.HashMap<String, Integer>(); m.put("x", 10); h.putAll(m); System.out.println(h.get("x"));"#);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn hashtable_clone_copies_entries() {
    let out = run_main(r#"java.util.Hashtable<String, Integer> h = new java.util.Hashtable<String, Integer>(); h.put("c", 3); java.util.Hashtable<String, Integer> c = (java.util.Hashtable<String, Integer>) h.clone(); System.out.println(c.get("c"));"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn hashtable_initial_capacity_constructor() {
    let out = run_main(r#"java.util.Hashtable<String, Integer> h = new java.util.Hashtable<String, Integer>(16); h.put("a", 1); System.out.println(h.size());"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn vector_synchronized_inherently_thread_safe_flag() {
    let out = run_main(r#"java.util.Vector<Integer> v = new java.util.Vector<Integer>(); v.add(1); System.out.println(v.size());"#);
    assert_eq!(out, vec!["1"]);
}
