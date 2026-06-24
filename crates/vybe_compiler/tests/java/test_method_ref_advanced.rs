use crate::helpers::{run_in_main, run_main};

#[test]
fn method_ref_static_parse_int_converts_string() {
    let out = run_main(
        "java.util.function.Function<String, Integer> parser = Integer::parseInt; System.out.println(parser.apply(\"42\"));",
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn method_ref_static_value_of_boxes_integer_as_string() {
    let out = run_main(
        "java.util.function.Function<Integer, String> fmt = String::valueOf; System.out.println(fmt.apply(15));",
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn method_ref_static_compare_to_on_strings() {
    let out = run_main(
        "java.util.function.BiFunction<String, String, Integer> cmp = String::compareTo; System.out.println(cmp.apply(\"a\", \"b\"));",
    );
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn method_ref_instance_length_on_string_receiver() {
    let out = run_main(
        "java.util.function.Function<String, Integer> len = String::length; System.out.println(len.apply(\"java\"));",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn method_ref_instance_to_upper_case_transforms_string() {
    let out = run_main(
        "java.util.function.Function<String, String> upper = String::toUpperCase; System.out.println(upper.apply(\"ok\"));",
    );
    assert_eq!(out, vec!["OK"]);
}

#[test]
fn method_ref_instance_bound_length_on_fixed_string() {
    let out = run_main(
        r#"String word = "four"; java.util.function.Supplier<Integer> len = word::length; System.out.println(len.get());"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn method_ref_instance_bound_to_upper_case_on_fixed_string() {
    let out = run_main(
        r#"String word = "hi"; java.util.function.Supplier<String> shout = word::toUpperCase; System.out.println(shout.get());"#,
    );
    assert_eq!(out, vec!["HI"]);
}

#[test]
fn method_ref_instance_bound_println_on_system_out() {
    let out = run_main(
        "java.util.function.Consumer<String> log = System.out::println; log.accept(\"bound\");",
    );
    assert_eq!(out, vec!["bound"]);
}

#[test]
fn method_ref_instance_bound_add_on_array_list() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); java.util.function.Consumer<Integer> adder = list::add; adder.accept(7); adder.accept(8); System.out.println(list.get(0)); System.out.println(list.size());",
    );
    assert_eq!(out, vec!["7", "2"]);
}

#[test]
fn method_ref_constructor_array_list_creates_empty_list() {
    let out = run_main(
        "java.util.function.Supplier<java.util.ArrayList<Integer>> factory = java.util.ArrayList::new; java.util.ArrayList<Integer> list = factory.get(); System.out.println(list.size()); System.out.println(list.isEmpty());",
    );
    assert_eq!(out, vec!["0", "true"]);
}

#[test]
fn method_ref_constructor_hash_set_creates_empty_set() {
    let out = run_main(
        "java.util.function.Supplier<java.util.HashSet<String>> factory = java.util.HashSet::new; java.util.HashSet<String> set = factory.get(); set.add(\"solo\"); System.out.println(set.size());",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn method_ref_constructor_hash_map_creates_empty_map() {
    let out = run_main(
        "java.util.function.Supplier<java.util.HashMap<String, Integer>> factory = java.util.HashMap::new; java.util.HashMap<String, Integer> map = factory.get(); map.put(\"k\", 1); System.out.println(map.get(\"k\"));",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn method_ref_array_constructor_creates_int_array_of_size() {
    let out = run_main(
        "java.util.function.IntFunction<int[]> maker = int[]::new; int[] arr = maker.apply(3); System.out.println(arr.length);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn method_ref_array_constructor_creates_string_array_of_size() {
    let out = run_main(
        "java.util.function.IntFunction<String[]> maker = String[]::new; String[] arr = maker.apply(2); System.out.println(arr.length);",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn method_ref_array_constructor_zero_length_array() {
    let out = run_main(
        "java.util.function.IntFunction<int[]> maker = int[]::new; int[] arr = maker.apply(0); System.out.println(arr.length);",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn method_ref_in_stream_map_applies_string_length() {
    let out = run_main(
        "java.util.List<Integer> lengths = java.util.Arrays.asList(\"a\", \"ab\", \"abc\").stream().map(String::length).toList(); System.out.println(lengths.get(0)); System.out.println(lengths.get(2));",
    );
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn method_ref_in_stream_map_applies_integer_parse_on_strings() {
    let out = run_main(
        "java.util.List<Integer> nums = java.util.Arrays.asList(\"1\", \"2\", \"3\").stream().map(Integer::parseInt).toList(); System.out.println(nums.get(0)); System.out.println(nums.get(2));",
    );
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn method_ref_in_stream_filter_uses_string_is_empty() {
    let out = run_main(
        "long count = java.util.Arrays.asList(\"\", \"a\", \"\").stream().filter(s -> !s.isEmpty()).count(); System.out.println(count);",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn method_ref_in_stream_map_to_int_uses_integer_int_value() {
    let out = run_main(
        "int sum = java.util.Arrays.asList(1, 2, 3).stream().mapToInt(Integer::intValue).sum(); System.out.println(sum);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn method_ref_in_stream_sorted_uses_natural_string_order() {
    let out = run_main(
        "java.util.List<String> sorted = java.util.Arrays.asList(\"c\", \"a\", \"b\").stream().sorted(String::compareTo).toList(); System.out.println(sorted.get(0)); System.out.println(sorted.get(2));",
    );
    assert_eq!(out, vec!["a", "c"]);
}

#[test]
fn method_ref_in_stream_for_each_prints_each_string() {
    let out = run_main(
        "java.util.Arrays.asList(\"x\", \"y\").stream().forEach(System.out::println);",
    );
    assert_eq!(out, vec!["x", "y"]);
}

#[test]
fn method_ref_in_stream_reduce_with_string_concat() {
    let out = run_main(
        "String joined = java.util.Arrays.asList(\"a\", \"b\").stream().reduce(\"\", String::concat); System.out.println(joined);",
    );
    assert_eq!(out, vec!["ab"]);
}

#[test]
fn method_ref_in_collectors_mapping_transforms_before_list() {
    let out = run_main(
        "java.util.List<Integer> lengths = java.util.Arrays.asList(\"a\", \"bb\").stream().collect(java.util.stream.Collectors.mapping(String::length, java.util.stream.Collectors.toList())); System.out.println(lengths.get(0)); System.out.println(lengths.get(1));",
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn method_ref_in_comparator_comparing_sorts_by_length() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"ccc\"); list.add(\"a\"); list.sort(java.util.Comparator.comparing(String::length)); System.out.println(list.get(0)); System.out.println(list.get(1));",
    );
    assert_eq!(out, vec!["a", "ccc"]);
}

#[test]
fn method_ref_list_for_each_with_println_reference() {
    let out = run_main(
        "java.util.Arrays.asList(1, 2, 3).forEach(System.out::println);",
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn method_ref_list_remove_if_with_string_is_empty() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"\"); list.add(\"x\"); list.removeIf(String::isEmpty); System.out.println(list.size()); System.out.println(list.get(0));",
    );
    assert_eq!(out, vec!["1", "x"]);
}

#[test]
fn method_ref_optional_map_applies_string_to_upper_case() {
    let out = run_main(
        r#"java.util.Optional<String> upper = java.util.Optional.of("ok").map(String::toUpperCase); System.out.println(upper.get());"#,
    );
    assert_eq!(out, vec!["OK"]);
}

#[test]
fn method_ref_optional_flat_map_wraps_non_empty_string() {
    let out = run_main(
        r#"java.util.Optional<String> out = java.util.Optional.of("7").flatMap(s -> java.util.Optional.of(String.valueOf(s.length()))); System.out.println(out.get());"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn method_ref_bi_function_integer_sum_via_method_handle_style() {
    let out = run_main(
        "java.util.function.BiFunction<Integer, Integer, Integer> add = Integer::sum; System.out.println(add.apply(4, 5));",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn method_ref_constructor_string_builder_then_append() {
    let out = run_main(
        "java.util.function.Supplier<StringBuilder> factory = StringBuilder::new; StringBuilder sb = factory.get(); sb.append(\"go\"); System.out.println(sb.toString());",
    );
    assert_eq!(out, vec!["go"]);
}

#[test]
fn method_ref_in_stream_distinct_on_mapped_strings() {
    let out = run_main(
        "long count = java.util.Arrays.asList(\"A\", \"a\", \"B\").stream().map(String::toLowerCase).distinct().count(); System.out.println(count);",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn method_ref_in_stream_flat_map_collection_stream() {
    let out = run_main(
        "long count = java.util.Arrays.asList(java.util.Arrays.asList(1, 2), java.util.Arrays.asList(3)).stream().flatMap(java.util.Collection::stream).count(); System.out.println(count);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn method_ref_static_math_abs_on_doubles_in_map() {
    let out = run_main(
        "java.util.List<Double> abs = java.util.Arrays.asList(-1.5, 2.5).stream().map(Math::abs).toList(); System.out.println(abs.get(0).intValue()); System.out.println(abs.get(1).intValue());",
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn method_ref_instance_bound_get_on_map_entry() {
    let out = run_in_main(
        "java.util.Map.Entry<String, Integer> e = java.util.Map.entry(\"k\", 9); java.util.function.Supplier<Integer> read = e::getValue; System.out.println(read.get());",
        "",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn method_ref_in_priority_queue_with_comparator_comparing() {
    let out = run_main(
        "java.util.PriorityQueue<String> pq = new java.util.PriorityQueue<String>(java.util.Comparator.comparing(String::length)); pq.offer(\"aaa\"); pq.offer(\"b\"); System.out.println(pq.poll());",
    );
    assert_eq!(out, vec!["b"]);
}

#[test]
fn method_ref_in_stream_max_using_comparator_comparing_length() {
    let out = run_main(
        "java.util.Optional<String> max = java.util.Arrays.asList(\"a\", \"ccc\", \"bb\").stream().max(java.util.Comparator.comparing(String::length)); System.out.println(max.get());",
    );
    assert_eq!(out, vec!["ccc"]);
}

#[test]
fn method_ref_array_constructor_in_stream_to_array() {
    let out = run_main(
        "String[] arr = java.util.Arrays.asList(\"a\", \"b\").stream().toArray(String[]::new); System.out.println(arr.length); System.out.println(arr[1]);",
    );
    assert_eq!(out, vec!["2", "b"]);
}

#[test]
fn method_ref_constructor_linked_list_populated_via_add() {
    let out = run_main(
        "java.util.function.Supplier<java.util.LinkedList<Integer>> factory = java.util.LinkedList::new; java.util.LinkedList<Integer> list = factory.get(); list.add(4); list.add(5); System.out.println(list.getFirst()); System.out.println(list.getLast());",
    );
    assert_eq!(out, vec!["4", "5"]);
}

#[test]
fn method_ref_in_stream_collect_joining_after_map_to_string() {
    let out = run_main(
        "String joined = java.util.Arrays.asList(1, 2, 3).stream().map(String::valueOf).collect(java.util.stream.Collectors.joining(\"\")); System.out.println(joined);",
    );
    assert_eq!(out, vec!["123"]);
}

#[test]
fn method_ref_static_objects_require_non_null_in_map() {
    let out = run_main(
        r#"java.util.List<String> out = java.util.Arrays.asList("a", "b").stream().map(java.util.Objects::requireNonNull).toList(); System.out.println(out.get(1));"#,
    );
    assert_eq!(out, vec!["b"]);
}

#[test]
fn method_ref_instance_is_present_on_optional_in_filter() {
    let out = run_main(
        r#"boolean any = java.util.Arrays.asList(java.util.Optional.of(1), java.util.Optional.empty()).stream().filter(java.util.Optional::isPresent).count() > 0; System.out.println(any);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn method_ref_in_tree_set_with_comparing_natural_order() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>(java.util.Comparator.naturalOrder()); set.add(3); set.add(1); set.add(2); System.out.println(set.first()); System.out.println(set.last());",
    );
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn method_ref_constructor_optional_empty_factory() {
    let out = run_main(
        "java.util.function.Supplier<java.util.Optional<String>> factory = java.util.Optional::empty; java.util.Optional<String> opt = factory.get(); System.out.println(opt.isPresent());",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn method_ref_in_stream_peek_with_println_reference() {
    let out = run_main(
        "java.util.List<Integer> out = java.util.Arrays.asList(1, 2).stream().peek(System.out::println).toList(); System.out.println(out.size());",
    );
    assert_eq!(out, vec!["1", "2", "2"]);
}

#[test]
fn method_ref_combines_static_and_instance_in_single_pipeline() {
    let out = run_main(
        "String result = java.util.Arrays.asList(\"  hi  \").stream().map(String::trim).map(String::toUpperCase).findFirst().get(); System.out.println(result);",
    );
    assert_eq!(out, vec!["HI"]);
}
