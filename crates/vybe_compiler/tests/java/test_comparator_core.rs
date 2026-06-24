use crate::helpers::{run_in_main, run_main};

#[test]
fn comparator_natural_order_sorts_integers_ascending() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(3); list.add(1); list.add(2); list.sort(java.util.Comparator.naturalOrder()); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn comparator_reverse_order_sorts_integers_descending() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(3); list.add(1); list.add(2); list.sort(java.util.Comparator.reverseOrder()); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["3", "1"]);
}

#[test]
fn comparator_natural_order_sorts_strings_lexicographically() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"cherry\"); list.add(\"apple\"); list.add(\"banana\"); list.sort(java.util.Comparator.naturalOrder()); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["apple", "cherry"]);
}

#[test]
fn comparator_reverse_order_sorts_strings_descending() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"cherry\"); list.add(\"apple\"); list.add(\"banana\"); list.sort(java.util.Comparator.reverseOrder()); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["cherry", "apple"]);
}

#[test]
fn comparator_comparing_by_string_length_orders_shortest_first() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"aaa\"); list.add(\"b\"); list.add(\"cc\"); list.sort(java.util.Comparator.comparing((String s) -> s.length())); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["b", "aaa"]);
}

#[test]
fn comparator_comparing_reversed_flips_length_order() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"aaa\"); list.add(\"b\"); list.add(\"cc\"); list.sort(java.util.Comparator.comparing((String s) -> s.length()).reversed()); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["aaa", "b"]);
}

#[test]
fn comparator_then_comparing_breaks_ties_lexicographically() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"bb\"); list.add(\"aa\"); list.add(\"cc\"); list.sort(java.util.Comparator.comparing((String s) -> s.length()).thenComparing((String s) -> s)); System.out.println(list.get(0)); System.out.println(list.get(1));",
    );
    assert_eq!(out, vec!["aa", "bb"]);
}

#[test]
fn comparator_then_comparing_chains_integer_then_string_keys() {
    let out = run_in_main(
        "java.util.ArrayList<Pair> list = new java.util.ArrayList<Pair>(); list.add(new Pair(2, \"b\")); list.add(new Pair(1, \"z\")); list.add(new Pair(1, \"a\")); list.sort(java.util.Comparator.comparing((Pair p) -> p.num).thenComparing((Pair p) -> p.label)); System.out.println(list.get(0).label); System.out.println(list.get(2).label);",
        "static class Pair { int num; String label; Pair(int num, String label) { this.num = num; this.label = label; } }",
    );
    assert_eq!(out, vec!["a", "b"]);
}

#[test]
fn comparator_comparing_integer_identity_sorts_naturally() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(9); list.add(2); list.add(5); list.sort(java.util.Comparator.comparing((Integer n) -> n)); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["2", "9"]);
}

#[test]
fn comparator_then_comparing_reversed_secondary_sorts_descending_ties() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"ab\"); list.add(\"ba\"); list.add(\"aa\"); list.sort(java.util.Comparator.comparing((String s) -> s.length()).thenComparing((String s) -> s).reversed()); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["ba", "aa"]);
}

#[test]
fn comparator_reversed_on_natural_order_restores_descending() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(3); list.add(2); list.sort(java.util.Comparator.naturalOrder().reversed()); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["3", "1"]);
}

#[test]
fn comparator_reversed_twice_matches_natural_order() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(3); list.add(1); list.add(2); list.sort(java.util.Comparator.reverseOrder().reversed()); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn comparator_collections_max_with_natural_order() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(4); list.add(9); list.add(2); System.out.println(java.util.Collections.max(list));",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn comparator_collections_min_with_natural_order() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(4); list.add(9); list.add(2); System.out.println(java.util.Collections.min(list));",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn comparator_collections_max_with_reverse_order() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(4); list.add(9); list.add(2); System.out.println(java.util.Collections.max(list, java.util.Comparator.reverseOrder()));",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn comparator_collections_min_with_reverse_order() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(4); list.add(9); list.add(2); System.out.println(java.util.Collections.min(list, java.util.Comparator.reverseOrder()));",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn comparator_comparing_by_absolute_value_orders_near_zero_first() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(-5); list.add(2); list.add(-1); list.sort(java.util.Comparator.comparing((Integer n) -> Math.abs(n))); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["-1", "-5"]);
}

#[test]
fn comparator_comparing_first_character_of_string() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"dog\"); list.add(\"ant\"); list.add(\"cat\"); list.sort(java.util.Comparator.comparing((String s) -> s.charAt(0))); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["ant", "dog"]);
}

#[test]
fn comparator_then_comparing_three_level_chain() {
    let out = run_in_main(
        "java.util.ArrayList<Triple> list = new java.util.ArrayList<Triple>(); list.add(new Triple(2, 1, \"c\")); list.add(new Triple(1, 2, \"b\")); list.add(new Triple(1, 1, \"a\")); list.sort(java.util.Comparator.comparing((Triple t) -> t.a).thenComparing((Triple t) -> t.b).thenComparing((Triple t) -> t.c)); System.out.println(list.get(0).c); System.out.println(list.get(2).c);",
        "static class Triple { int a; int b; String c; Triple(int a, int b, String c) { this.a = a; this.b = b; this.c = c; } }",
    );
    assert_eq!(out, vec!["a", "c"]);
}

#[test]
fn comparator_arrays_sort_with_natural_order() {
    let out = run_main(
        "Integer[] arr = new Integer[] {3, 1, 2}; java.util.Arrays.sort(arr, java.util.Comparator.naturalOrder()); System.out.println(arr[0]); System.out.println(arr[2]);",
    );
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn comparator_arrays_sort_with_reverse_order() {
    let out = run_main(
        "Integer[] arr = new Integer[] {3, 1, 2}; java.util.Arrays.sort(arr, java.util.Comparator.reverseOrder()); System.out.println(arr[0]); System.out.println(arr[2]);",
    );
    assert_eq!(out, vec!["3", "1"]);
}

#[test]
fn comparator_treeset_with_reverse_order_iterates_descending() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(java.util.Comparator.reverseOrder()); set.add(1); set.add(3); set.add(2); System.out.println(set.first()); System.out.println(set.last());",
    );
    assert_eq!(out, vec!["3", "1"]);
}

#[test]
fn comparator_treemap_with_comparing_key_length() {
    let out = run_main(
        "java.util.TreeMap<String, Integer> map = new java.util.TreeMap<String, Integer>(java.util.Comparator.comparing((String s) -> s.length())); map.put(\"long\", 2); map.put(\"a\", 1); map.put(\"mid\", 3); System.out.println(map.firstKey()); System.out.println(map.lastKey());",
    );
    assert_eq!(out, vec!["a", "long"]);
}

#[test]
fn comparator_priorityqueue_with_reverse_order_polls_largest() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>(java.util.Comparator.reverseOrder()); pq.offer(2); pq.offer(8); pq.offer(5); System.out.println(pq.poll());",
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn comparator_comparing_method_reference_string_case() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"b\"); list.add(\"a\"); list.add(\"c\"); list.sort(java.util.Comparator.comparing((String s) -> s)); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["a", "c"]);
}

#[test]
fn comparator_then_comparing_int_modulo_groups_by_remainder() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(7); list.add(4); list.add(9); list.add(2); list.sort(java.util.Comparator.comparing((Integer n) -> n % 5).thenComparing((Integer n) -> n)); System.out.println(list.get(0)); System.out.println(list.get(3));",
    );
    assert_eq!(out, vec!["4", "9"]);
}

#[test]
fn comparator_lambda_sort_orders_descending_without_static_factory() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(3); list.add(2); list.sort((a, b) -> b - a); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["3", "1"]);
}

#[test]
fn comparator_anonymous_class_sorts_strings_by_length() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"long\"); list.add(\"a\"); list.add(\"mid\"); list.sort(new java.util.Comparator<String>() { public int compare(String a, String b) { return Integer.compare(a.length(), b.length()); } }); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["a", "long"]);
}

#[test]
fn comparator_collections_sort_with_comparing_length() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"zzz\"); list.add(\"a\"); list.add(\"bb\"); java.util.Collections.sort(list, java.util.Comparator.comparing((String s) -> s.length())); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["a", "zzz"]);
}

#[test]
fn comparator_comparing_double_field_via_then_comparing() {
    let out = run_in_main(
        "java.util.ArrayList<Score> list = new java.util.ArrayList<Score>(); list.add(new Score(2.5)); list.add(new Score(1.1)); list.add(new Score(3.0)); list.sort(java.util.Comparator.comparing((Score s) -> s.value)); System.out.println(list.get(0).value); System.out.println(list.get(2).value);",
        "static class Score { double value; Score(double value) { this.value = value; } }",
    );
    assert_eq!(out, vec!["1.1", "3.0"]);
}

#[test]
fn comparator_natural_order_on_empty_list_is_noop() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.sort(java.util.Comparator.naturalOrder()); System.out.println(list.size());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn comparator_comparing_then_comparing_equal_primary_preserves_secondary() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"bx\"); list.add(\"ax\"); list.add(\"ay\"); list.sort(java.util.Comparator.comparing((String s) -> s.substring(1)).thenComparing((String s) -> s.charAt(0))); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["ax", "bx"]);
}

#[test]
fn comparator_reverse_order_on_single_element_list() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(42); list.sort(java.util.Comparator.reverseOrder()); System.out.println(list.get(0));",
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn comparator_comparing_integer_string_value_of_sorts_numeric_strings() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"10\"); list.add(\"2\"); list.add(\"1\"); list.sort(java.util.Comparator.comparing((String s) -> Integer.parseInt(s))); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["1", "10"]);
}

#[test]
fn comparator_then_comparing_reversed_on_secondary_key_only() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"ab\"); list.add(\"aa\"); list.add(\"ac\"); list.sort(java.util.Comparator.comparing((String s) -> s.length()).thenComparing((String s) -> s).reversed()); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["ac", "aa"]);
}

#[test]
fn comparator_treeset_natural_order_first_is_minimum() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(java.util.Comparator.naturalOrder()); set.add(5); set.add(1); set.add(3); System.out.println(set.first());",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn comparator_treemap_reverse_order_first_is_maximum_key() {
    let out = run_main(
        "java.util.TreeMap<Integer, String> map = new java.util.TreeMap<Integer, String>(java.util.Comparator.reverseOrder()); map.put(1, \"a\"); map.put(3, \"c\"); map.put(2, \"b\"); System.out.println(map.firstKey());",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn comparator_priorityqueue_comparing_polls_lowest_priority_field() {
    let out = run_in_main(
        "java.util.PriorityQueue<Job> pq = new java.util.PriorityQueue<Job>(java.util.Comparator.comparing((Job j) -> j.rank)); pq.offer(new Job(30)); pq.offer(new Job(5)); pq.offer(new Job(15)); System.out.println(pq.poll().rank); System.out.println(pq.poll().rank);",
        "static class Job { int rank; Job(int rank) { this.rank = rank; } }",
    );
    assert_eq!(out, vec!["5", "15"]);
}

#[test]
fn comparator_chaining_then_comparing_after_reversed_primary() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"aab\"); list.add(\"aaa\"); list.add(\"baa\"); list.sort(java.util.Comparator.comparing((String s) -> s.length()).reversed().thenComparing((String s) -> s)); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["baa", "aaa"]);
}

#[test]
fn comparator_comparing_identity_equals_natural_order_on_integers() {
    let out = run_main(
        "java.util.ArrayList<Integer> natural = new java.util.ArrayList<Integer>(); natural.add(3); natural.add(1); natural.sort(java.util.Comparator.naturalOrder()); java.util.ArrayList<Integer> comparing = new java.util.ArrayList<Integer>(); comparing.add(3); comparing.add(1); comparing.sort(java.util.Comparator.comparing((Integer n) -> n)); System.out.println(natural.get(0)); System.out.println(comparing.get(0));",
    );
    assert_eq!(out, vec!["1", "1"]);
}
