use crate::helpers::run_main;

#[test]
fn stack_default_constructor_empty() {
    let out = run_main(r#"java.util.Stack<Integer> s = new java.util.Stack<Integer>(); System.out.println(s.isEmpty());"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn stack_push_adds_element_on_top() {
    let out = run_main(r#"java.util.Stack<String> s = new java.util.Stack<String>(); s.push("top"); System.out.println(s.peek());"#);
    assert_eq!(out, vec!["top"]);
}

#[test]
fn stack_push_returns_pushed_element() {
    let out = run_main(r#"java.util.Stack<Integer> s = new java.util.Stack<Integer>(); System.out.println(s.push(42));"#);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn stack_pop_removes_top_element() {
    let out = run_main(r#"java.util.Stack<Integer> s = new java.util.Stack<Integer>(); s.push(1); s.push(2); System.out.println(s.pop());"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn stack_peek_reads_without_removing() {
    let out = run_main(r#"java.util.Stack<String> s = new java.util.Stack<String>(); s.push("a"); s.peek(); System.out.println(s.peek());"#);
    assert_eq!(out, vec!["a"]);
}

#[test]
fn stack_peek_after_pop_returns_new_top() {
    let out = run_main(r#"java.util.Stack<Integer> s = new java.util.Stack<Integer>(); s.push(10); s.push(20); s.pop(); System.out.println(s.peek());"#);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn stack_empty_false_after_push() {
    let out = run_main(r#"java.util.Stack<Integer> s = new java.util.Stack<Integer>(); s.push(1); System.out.println(s.isEmpty());"#);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn stack_empty_true_after_popping_all() {
    let out = run_main(r#"java.util.Stack<Integer> s = new java.util.Stack<Integer>(); s.push(1); s.pop(); System.out.println(s.isEmpty());"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn stack_search_returns_one_based_index_from_top() {
    let out = run_main(r#"java.util.Stack<String> s = new java.util.Stack<String>(); s.push("a"); s.push("b"); s.push("c"); System.out.println(s.search("c"));"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn stack_search_finds_element_below_top() {
    let out = run_main(r#"java.util.Stack<String> s = new java.util.Stack<String>(); s.push("a"); s.push("b"); s.push("c"); System.out.println(s.search("a"));"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn stack_search_returns_negative_when_absent() {
    let out = run_main(r#"java.util.Stack<Integer> s = new java.util.Stack<Integer>(); s.push(1); System.out.println(s.search(99));"#);
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn stack_lifo_order_with_three_pushes_and_pops() {
    let out = run_main(r#"java.util.Stack<Integer> s = new java.util.Stack<Integer>(); s.push(1); s.push(2); s.push(3); System.out.println(s.pop()); System.out.println(s.pop()); System.out.println(s.pop());"#);
    assert_eq!(out, vec!["3", "2", "1"]);
}

#[test]
fn stack_size_grows_with_push() {
    let out = run_main(r#"java.util.Stack<Integer> s = new java.util.Stack<Integer>(); s.push(1); s.push(2); System.out.println(s.size());"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn stack_size_shrinks_with_pop() {
    let out = run_main(r#"java.util.Stack<Integer> s = new java.util.Stack<Integer>(); s.push(1); s.push(2); s.pop(); System.out.println(s.size());"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn stack_element_at_zero_is_bottom() {
    let out = run_main(r#"java.util.Stack<String> s = new java.util.Stack<String>(); s.push("bottom"); s.push("top"); System.out.println(s.elementAt(0));"#);
    assert_eq!(out, vec!["bottom"]);
}

#[test]
fn stack_element_at_top_index() {
    let out = run_main(r#"java.util.Stack<Integer> s = new java.util.Stack<Integer>(); s.push(10); s.push(20); System.out.println(s.elementAt(1));"#);
    assert_eq!(out, vec!["20"]);
}

#[test]
fn stack_first_element_is_bottom() {
    let out = run_main(r#"java.util.Stack<String> s = new java.util.Stack<String>(); s.push("first"); s.push("second"); System.out.println(s.firstElement());"#);
    assert_eq!(out, vec!["first"]);
}

#[test]
fn stack_last_element_is_top() {
    let out = run_main(r#"java.util.Stack<String> s = new java.util.Stack<String>(); s.push("first"); s.push("last"); System.out.println(s.lastElement());"#);
    assert_eq!(out, vec!["last"]);
}

#[test]
fn stack_push_null_element_allowed() {
    let out = run_main(r#"java.util.Stack<String> s = new java.util.Stack<String>(); s.push(null); System.out.println(s.peek() == null);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn stack_clear_empties_all_elements() {
    let out = run_main(r#"java.util.Stack<Integer> s = new java.util.Stack<Integer>(); s.push(1); s.push(2); s.clear(); System.out.println(s.isEmpty());"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn stack_contains_finds_pushed_element() {
    let out = run_main(r#"java.util.Stack<String> s = new java.util.Stack<String>(); s.push("find"); System.out.println(s.contains("find"));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn stack_contains_false_for_absent() {
    let out = run_main(r#"java.util.Stack<Integer> s = new java.util.Stack<Integer>(); s.push(1); System.out.println(s.contains(2));"#);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn stack_iterator_traverses_bottom_to_top() {
    let out = run_main(r#"java.util.Stack<Integer> s = new java.util.Stack<Integer>(); s.push(1); s.push(2); java.util.Iterator<Integer> it = s.iterator(); System.out.println(it.next());"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn stack_add_method_same_as_push() {
    let out = run_main(r#"java.util.Stack<Integer> s = new java.util.Stack<Integer>(); s.add(5); System.out.println(s.peek());"#);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn stack_remove_top_by_index() {
    let out = run_main(r#"java.util.Stack<String> s = new java.util.Stack<String>(); s.push("a"); s.push("b"); s.remove(1); System.out.println(s.peek());"#);
    assert_eq!(out, vec!["a"]);
}

#[test]
fn stack_set_replaces_element_at_index() {
    let out = run_main(r#"java.util.Stack<Integer> s = new java.util.Stack<Integer>(); s.push(1); s.push(2); s.set(0, 9); System.out.println(s.elementAt(0));"#);
    assert_eq!(out, vec!["9"]);
}

#[test]
fn stack_to_array_length_matches_size() {
    let out = run_main(r#"java.util.Stack<Integer> s = new java.util.Stack<Integer>(); s.push(1); s.push(2); System.out.println(s.toArray().length);"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn stack_clone_copies_elements() {
    let out = run_main(r#"java.util.Stack<Integer> s = new java.util.Stack<Integer>(); s.push(7); java.util.Stack<Integer> c = (java.util.Stack<Integer>) s.clone(); System.out.println(c.peek());"#);
    assert_eq!(out, vec!["7"]);
}
