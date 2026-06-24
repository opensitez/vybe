use crate::helpers::run_main;

#[test]
fn arraylist_add_and_get_preserve_order() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<>(); list.add(10); list.add(20); System.out.println(list.get(0)); System.out.println(list.get(1));",
    );
    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn arraylist_size_counts_elements() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<>(); list.add(\"a\"); list.add(\"b\"); list.add(\"c\"); System.out.println(list.size());",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn hashmap_put_get_roundtrip() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<>(); map.put(\"x\", 9); System.out.println(map.get(\"x\"));",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn hashset_rejects_duplicate_entries() {
    let out = run_main(
        "java.util.HashSet<Integer> set = new java.util.HashSet<>(); set.add(1); set.add(1); set.add(2); System.out.println(set.size());",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn linkedlist_remove_first_element() {
    let out = run_main(
        "java.util.LinkedList<Integer> list = new java.util.LinkedList<>(); list.add(5); list.add(6); list.removeFirst(); System.out.println(list.get(0));",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn collections_sort_orders_integers() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<>(); list.add(3); list.add(1); list.add(2); java.util.Collections.sort(list); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["1", "3"]);
}
