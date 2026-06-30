use crate::helpers::run_main;

#[test]
fn collectors_reducing_sum_with_identity() {
    let out = run_main(
        "Integer s = java.util.Arrays.asList(1, 2, 3).stream().collect(java.util.stream.Collectors.reducing(0, (a, b) -> a + b)); System.out.println(s);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn collectors_reducing_product_no_identity() {
    let out = run_main(
        "java.util.Optional<Integer> p = java.util.Arrays.asList(2, 3, 4).stream().collect(java.util.stream.Collectors.reducing((a, b) -> a * b)); System.out.println(p.get());",
    );
    assert_eq!(out, vec!["24"]);
}

#[test]
fn collectors_reducing_max_via_binary_op() {
    let out = run_main(
        "java.util.Optional<Integer> m = java.util.Arrays.asList(3, 9, 5).stream().collect(java.util.stream.Collectors.reducing((a, b) -> a > b ? a : b)); System.out.println(m.get());",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn collectors_reducing_empty_yields_empty_optional() {
    let out = run_main(
        "boolean present = java.util.Arrays.asList(1).stream().filter(n -> n > 5).collect(java.util.stream.Collectors.reducing((a, b) -> a + b)).isPresent(); System.out.println(present);",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn collectors_reducing_with_mapping() {
    let out = run_main(
        "Integer s = java.util.Arrays.asList(\"1\", \"2\", \"3\").stream().collect(java.util.stream.Collectors.reducing(0, s -> Integer.parseInt(s), (a, b) -> a + b)); System.out.println(s);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn collectors_max_by_natural_order() {
    let out = run_main(
        "java.util.Optional<Integer> m = java.util.Arrays.asList(3, 9, 5).stream().collect(java.util.stream.Collectors.maxBy(Integer::compareTo)); System.out.println(m.get());",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn collectors_min_by_natural_order() {
    let out = run_main(
        "java.util.Optional<Integer> m = java.util.Arrays.asList(3, 9, 5).stream().collect(java.util.stream.Collectors.minBy(Integer::compareTo)); System.out.println(m.get());",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn collectors_max_by_custom_comparator() {
    let out = run_main(
        "java.util.Optional<String> m = java.util.Arrays.asList(\"a\", \"ccc\", \"bb\").stream().collect(java.util.stream.Collectors.maxBy((a, b) -> a.length() - b.length())); System.out.println(m.get());",
    );
    assert_eq!(out, vec!["ccc"]);
}

#[test]
fn collectors_min_by_string_length() {
    let out = run_main(
        "java.util.Optional<String> m = java.util.Arrays.asList(\"a\", \"ccc\", \"bb\").stream().collect(java.util.stream.Collectors.minBy((a, b) -> a.length() - b.length())); System.out.println(m.get());",
    );
    assert_eq!(out, vec!["a"]);
}

#[test]
fn collectors_max_by_empty_stream() {
    let out = run_main(
        "boolean present = java.util.Arrays.<Integer>asList().stream().collect(java.util.stream.Collectors.maxBy(Integer::compareTo)).isPresent(); System.out.println(present);",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn collectors_to_collection_array_list() {
    let out = run_main(
        "int sz = java.util.Arrays.asList(1, 2, 3).stream().collect(java.util.stream.Collectors.toCollection(java.util.ArrayList::new)).size(); System.out.println(sz);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn collectors_to_collection_hash_set() {
    let out = run_main(
        "int sz = java.util.Arrays.asList(1, 2, 2).stream().collect(java.util.stream.Collectors.toCollection(java.util.HashSet::new)).size(); System.out.println(sz);",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn collectors_to_unmodifiable_list() {
    let out = run_main(
        "java.util.List<Integer> list = java.util.Arrays.asList(1, 2).stream().collect(java.util.stream.Collectors.toUnmodifiableList()); System.out.println(list.size());",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn collectors_to_unmodifiable_set() {
    let out = run_main(
        "int sz = java.util.Arrays.asList(1, 2, 2).stream().collect(java.util.stream.Collectors.toUnmodifiableSet()).size(); System.out.println(sz);",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn collectors_to_unmodifiable_map() {
    let out = run_main(
        "java.util.Map<Integer, Integer> map = java.util.Arrays.asList(1, 2).stream().collect(java.util.stream.Collectors.toUnmodifiableMap(n -> n, n -> n * 10)); System.out.println(map.get(2));",
    );
    assert_eq!(out, vec!["20"]);
}

#[test]
fn collectors_teeing_sum_and_count() {
    let out = run_main(
        "Long r = java.util.Arrays.asList(1, 2, 3).stream().collect(java.util.stream.Collectors.teeing(java.util.stream.Collectors.summingInt(n -> n), java.util.stream.Collectors.counting(), (s, c) -> (long) s + c)); System.out.println(r);",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn collectors_collecting_and_then_modify() {
    let out = run_main(
        "java.util.List<Integer> list = java.util.Arrays.asList(1, 2, 3).stream().collect(java.util.stream.Collectors.collectingAndThen(java.util.stream.Collectors.toList(), l -> { l.add(99); return l; })); System.out.println(list.get(3));",
    );
    assert_eq!(out, vec!["99"]);
}

#[test]
fn collectors_summing_double() {
    let out = run_main(
        "Double s = java.util.Arrays.asList(1.5, 2.5).stream().collect(java.util.stream.Collectors.summingDouble(n -> n)); System.out.println((int) (double) s);",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn collectors_summing_long() {
    let out = run_main(
        "Long s = java.util.Arrays.asList(1L, 2L, 3L).stream().collect(java.util.stream.Collectors.summingLong(n -> (long) n)); System.out.println(s);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn collectors_averaging_double() {
    let out = run_main(
        "Double a = java.util.Arrays.asList(2.0, 4.0).stream().collect(java.util.stream.Collectors.averagingDouble(n -> n)); System.out.println((int) (double) a);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn collectors_averaging_long() {
    let out = run_main(
        "Double a = java.util.Arrays.asList(2L, 4L, 6L).stream().collect(java.util.stream.Collectors.averagingLong(n -> n)); System.out.println(a.intValue());",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn collectors_joining_with_prefix_suffix() {
    let out = run_main(
        "String j = java.util.Arrays.asList(\"a\", \"b\").stream().collect(java.util.stream.Collectors.joining(\",\", \"[\", \"]\")); System.out.println(j);",
    );
    assert_eq!(out, vec!["[a,b]"]);
}

#[test]
fn collectors_mapping_to_int_sum() {
    let out = run_main(
        "Integer s = java.util.Arrays.asList(\"1\", \"2\").stream().collect(java.util.stream.Collectors.mapping(Integer::parseInt, java.util.stream.Collectors.summingInt(n -> n))); System.out.println(s);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn collectors_flat_mapping_splits() {
    let out = run_main(
        "java.util.List<String> out = java.util.Arrays.asList(\"a b\").stream().collect(java.util.stream.Collectors.flatMapping(s -> java.util.Arrays.stream(s.split(\" \")), java.util.stream.Collectors.toList())); System.out.println(out.size());",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn collectors_grouping_by_concurrent_hash_map() {
    let out = run_main(
        "int sz = java.util.Arrays.asList(1, 2, 3).stream().collect(java.util.stream.Collectors.groupingBy(n -> n % 2, java.util.concurrent.ConcurrentHashMap::new, java.util.stream.Collectors.toList())).size(); System.out.println(sz);",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn collectors_partitioning_by_to_set() {
    let out = run_main(
        "java.util.Map<Boolean, java.util.Set<Integer>> parts = java.util.Arrays.asList(1, 2, 3, 4).stream().collect(java.util.stream.Collectors.partitioningBy(n -> n % 2 == 0, java.util.stream.Collectors.toSet())); System.out.println(parts.get(true).size());",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn collectors_reducing_string_concat() {
    let out = run_main(
        "String s = java.util.Arrays.asList(\"a\", \"b\", \"c\").stream().collect(java.util.stream.Collectors.reducing(\"\", (a, b) -> a + b)); System.out.println(s);",
    );
    assert_eq!(out, vec!["abc"]);
}

#[test]
fn collectors_min_by_after_filter() {
    let out = run_main(
        "java.util.Optional<Integer> m = java.util.Arrays.asList(1, 5, 3, 8).stream().filter(n -> n < 6).collect(java.util.stream.Collectors.minBy(Integer::compareTo)); System.out.println(m.get());",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn collectors_max_by_after_map() {
    let out = run_main(
        "java.util.Optional<Integer> m = java.util.Arrays.asList(\"1\", \"9\", \"5\").stream().map(Integer::parseInt).collect(java.util.stream.Collectors.maxBy(Integer::compareTo)); System.out.println(m.get());",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn collectors_to_collection_linked_list_order() {
    let out = run_main(
        "String first = java.util.Arrays.asList(\"b\", \"a\").stream().collect(java.util.stream.Collectors.toCollection(java.util.LinkedList::new)).getFirst(); System.out.println(first);",
    );
    assert_eq!(out, vec!["b"]);
}

#[test]
fn collectors_teeing_join_both_strings() {
    let out = run_main(
        "String r = java.util.Arrays.asList(\"x\", \"y\").stream().collect(java.util.stream.Collectors.teeing(java.util.stream.Collectors.joining(), java.util.stream.Collectors.counting(), (a, b) -> a + b)); System.out.println(r);",
    );
    assert_eq!(out, vec!["xy2"]);
}

#[test]
fn collectors_collecting_and_then_list_size() {
    let out = run_main(
        "Integer sz = java.util.Arrays.asList(1, 2, 3).stream().collect(java.util.stream.Collectors.collectingAndThen(java.util.stream.Collectors.counting(), Long::intValue)); System.out.println(sz);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn collectors_reducing_singleton_optional() {
    let out = run_main(
        "java.util.Optional<Integer> v = java.util.Arrays.asList(42).stream().collect(java.util.stream.Collectors.reducing((a, b) -> a + b)); System.out.println(v.get());",
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn collectors_summing_int_with_mapper_square() {
    let out = run_main(
        "Integer s = java.util.Arrays.asList(1, 2, 3).stream().collect(java.util.stream.Collectors.summingInt(n -> n * n)); System.out.println(s);",
    );
    assert_eq!(out, vec!["14"]);
}

#[test]
fn collectors_averaging_int_empty_is_nan() {
    let out = run_main(
        "Double a = java.util.Arrays.<Integer>asList().stream().collect(java.util.stream.Collectors.averagingInt(n -> n)); System.out.println(a.isNaN());",
    );
    assert_eq!(out, vec!["true"]);
}

