use crate::helpers::{run_in_main, run_main};

#[test]
fn lambda_captures_single_effectively_final_local() {
    let out = run_main(
        "int base = 10; java.util.function.IntSupplier supplier = () -> base + 1; System.out.println(supplier.getAsInt());",
    );
    assert_eq!(out, vec!["11"]);
}

#[test]
fn lambda_captures_two_effectively_final_locals() {
    let out = run_main(
        "int a = 2; int b = 3; java.util.function.Supplier<Integer> sum = () -> a + b; System.out.println(sum.get());",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn lambda_captures_string_reference_read_only() {
    let out = run_main(
        r#"String label = "hi"; java.util.function.Supplier<String> read = () -> label; System.out.println(read.get());"#,
    );
    assert_eq!(out, vec!["hi"]);
}

#[test]
fn lambda_captures_wrapper_integer_value() {
    let out = run_main(
        "Integer boxed = 7; java.util.function.Supplier<Integer> read = () -> boxed; System.out.println(read.get());",
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn lambda_captures_array_element_via_box_mutation() {
    let out = run_main(
        "int[] box = new int[]{0}; java.util.function.IntSupplier next = () -> { box[0] = box[0] + 1; return box[0]; }; System.out.println(next.getAsInt()); System.out.println(next.getAsInt());",
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn lambda_captures_outer_loop_counter_after_loop_finishes() {
    let out = run_main(
        "int last = 0; for (int i = 1; i <= 3; i++) { last = i; } java.util.function.Supplier<Integer> read = () -> last; System.out.println(read.get());",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn lambda_captures_final_local_int_for_arithmetic() {
    let out = run_main(
        "final int seed = 6; java.util.function.IntUnaryOperator bump = n -> n + seed; System.out.println(bump.applyAsInt(4));",
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn lambda_in_list_sort_orders_integers_ascending() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(3); list.add(1); list.add(2); list.sort((a, b) -> a - b); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn lambda_in_list_sort_orders_integers_descending() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(3); list.add(2); list.sort((a, b) -> b - a); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["3", "1"]);
}

#[test]
fn lambda_in_collections_sort_orders_strings_lexicographically() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"cherry\"); list.add(\"apple\"); java.util.Collections.sort(list, (a, b) -> a.compareTo(b)); System.out.println(list.get(0));",
    );
    assert_eq!(out, vec!["apple"]);
}

#[test]
fn lambda_in_stream_sorted_orders_mapped_lengths() {
    let out = run_main(
        "java.util.List<String> sorted = java.util.Arrays.asList(\"ccc\", \"a\", \"bb\").stream().sorted((a, b) -> a.length() - b.length()).toList(); System.out.println(sorted.get(0)); System.out.println(sorted.get(2));",
    );
    assert_eq!(out, vec!["a", "ccc"]);
}

#[test]
fn lambda_in_priority_queue_orders_by_custom_comparator() {
    let out = run_main(
        "java.util.PriorityQueue<Integer> pq = new java.util.PriorityQueue<Integer>((a, b) -> b - a); pq.offer(1); pq.offer(3); pq.offer(2); System.out.println(pq.poll()); System.out.println(pq.poll());",
    );
    assert_eq!(out, vec!["3", "2"]);
}

#[test]
fn lambda_returning_lambda_applies_outer_parameter_twice() {
    let out = run_main(
        "java.util.function.IntUnaryOperator doubler = x -> x * 2; java.util.function.Function<Integer, java.util.function.IntUnaryOperator> factory = base -> n -> base + n; java.util.function.IntUnaryOperator addBase = factory.apply(10); System.out.println(addBase.applyAsInt(5));",
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn lambda_returning_lambda_creates_multiplier_closure() {
    let out = run_main(
        "java.util.function.Function<Integer, java.util.function.IntUnaryOperator> scaleBy = factor -> n -> n * factor; java.util.function.IntUnaryOperator triple = scaleBy.apply(3); System.out.println(triple.applyAsInt(4));",
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn lambda_returning_lambda_string_prefix_factory() {
    let out = run_main(
        r#"java.util.function.Function<String, java.util.function.Function<String, java.util.function.Function<String, String>>> joiner = sep -> a -> b -> a + sep + b; System.out.println(joiner.apply("-").apply("x").apply("y"));"#,
    );
    assert_eq!(out, vec!["x-y"]);
}

#[test]
fn lambda_returning_lambda_predicate_negates_inner_test() {
    let out = run_main(
        "java.util.function.Function<java.util.function.Predicate<Integer>, java.util.function.Predicate<Integer>> negate = inner -> n -> !inner.test(n); java.util.function.Predicate<Integer> positive = n -> n > 0; java.util.function.Predicate<Integer> nonPositive = negate.apply(positive); System.out.println(nonPositive.test(-1)); System.out.println(nonPositive.test(2));",
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn comparator_comparing_by_string_length_orders_ascending() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"ccc\"); list.add(\"a\"); list.add(\"bb\"); list.sort(java.util.Comparator.comparing(s -> s.length())); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["a", "ccc"]);
}

#[test]
fn comparator_comparing_by_integer_identity_sorts_numbers() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(30); list.add(10); list.add(20); list.sort(java.util.Comparator.comparing(n -> n)); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["10", "30"]);
}

#[test]
fn comparator_comparing_with_method_reference_on_strings() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"b\"); list.add(\"a\"); list.sort(java.util.Comparator.comparing(String::toString)); System.out.println(list.get(0)); System.out.println(list.get(1));",
    );
    assert_eq!(out, vec!["a", "b"]);
}

#[test]
fn comparator_then_comparing_breaks_ties_by_second_key() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"bb\"); list.add(\"aa\"); list.add(\"ab\"); list.sort(java.util.Comparator.comparing((String s) -> s.length()).thenComparing(s -> s)); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["aa", "bb"]);
}

#[test]
fn comparator_then_comparing_orders_by_length_then_lexicographic() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"b\"); list.add(\"a\"); list.add(\"cc\"); list.sort(java.util.Comparator.comparing((String s) -> s.length()).thenComparing(s -> s)); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["a", "cc"]);
}

#[test]
fn comparator_comparing_int_reversed_orders_descending() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(3); list.add(2); list.sort(java.util.Comparator.comparing((Integer n) -> n).reversed()); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["3", "1"]);
}

#[test]
fn comparator_natural_order_sorts_integers_ascending() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(5); list.add(1); list.add(4); list.sort(java.util.Comparator.naturalOrder()); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["1", "5"]);
}

#[test]
fn comparator_reverse_order_sorts_integers_descending() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(5); list.add(1); list.add(4); list.sort(java.util.Comparator.reverseOrder()); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["5", "1"]);
}

#[test]
fn lambda_captures_multiple_types_in_single_expression() {
    let out = run_main(
        r#"int n = 4; String tag = "x"; java.util.function.Supplier<String> fmt = () -> tag + n; System.out.println(fmt.get());"#,
    );
    assert_eq!(out, vec!["x4"]);
}

#[test]
fn lambda_used_as_comparator_in_tree_set_orders_elements() {
    let out = run_main(
        "java.util.TreeSet<Integer> set = new java.util.TreeSet<Integer>((a, b) -> b - a); set.add(1); set.add(3); set.add(2); System.out.println(set.first()); System.out.println(set.last());",
    );
    assert_eq!(out, vec!["3", "1"]);
}

#[test]
fn lambda_in_stream_max_with_comparator_finds_custom_maximum() {
    let out = run_main(
        "java.util.Optional<String> max = java.util.Arrays.asList(\"a\", \"ccc\", \"bb\").stream().max((a, b) -> a.length() - b.length()); System.out.println(max.get());",
    );
    assert_eq!(out, vec!["ccc"]);
}

#[test]
fn lambda_in_stream_min_with_comparator_finds_custom_minimum() {
    let out = run_main(
        "java.util.Optional<String> min = java.util.Arrays.asList(\"ccc\", \"a\", \"bb\").stream().min((a, b) -> a.length() - b.length()); System.out.println(min.get());",
    );
    assert_eq!(out, vec!["a"]);
}

#[test]
fn lambda_block_body_captures_local_and_returns_sum() {
    let out = run_main(
        "int offset = 5; java.util.function.Function<Integer, Integer> addOffset = n -> { return n + offset; }; System.out.println(addOffset.apply(7));",
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn lambda_nested_inside_another_lambda_invocation() {
    let out = run_main(
        "java.util.function.Function<Integer, Integer> twice = n -> n * 2; java.util.function.Function<Integer, Integer> compose = n -> twice.apply(twice.apply(n)); System.out.println(compose.apply(3));",
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn lambda_in_foreach_with_captured_counter_array() {
    let out = run_main(
        "int[] count = new int[]{0}; java.util.Arrays.asList(1, 2, 3).forEach(n -> { count[0] = count[0] + 1; }); System.out.println(count[0]);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn lambda_in_remove_if_on_list_filters_in_place() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); list.add(3); list.removeIf(n -> n % 2 == 0); System.out.println(list.size()); System.out.println(list.get(1));",
    );
    assert_eq!(out, vec!["2", "3"]);
}

#[test]
fn lambda_in_replace_all_transforms_list_elements() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); list.replaceAll(n -> n * 10); System.out.println(list.get(0)); System.out.println(list.get(1));",
    );
    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn comparator_then_comparing_int_orders_by_length_then_char() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"ba\"); list.add(\"ab\"); list.add(\"a\"); list.sort(java.util.Comparator.comparing((String s) -> s.length()).thenComparingInt(s -> s.charAt(0))); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["a", "ba"]);
}

#[test]
fn lambda_captures_static_final_field_from_helper() {
    let out = run_in_main(
        "java.util.function.IntSupplier read = () -> OFFSET; System.out.println(read.getAsInt());",
        "static final int OFFSET = 9;",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn lambda_in_map_merge_function_combines_colliding_values() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put(\"a\", 1); map.merge(\"a\", 2, (oldV, newV) -> oldV + newV); System.out.println(map.get(\"a\"));",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn lambda_in_compute_if_absent_initializes_missing_key() {
    let out = run_main(
        "java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.computeIfAbsent(\"k\", key -> key.length()); System.out.println(map.get(\"k\"));",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn lambda_returning_lambda_curries_addition() {
    let out = run_main(
        "java.util.function.Function<Integer, java.util.function.Function<Integer, Integer>> add = a -> b -> a + b; System.out.println(add.apply(2).apply(5));",
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn comparator_comparing_nulls_last_places_null_at_end() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"b\"); list.add(null); list.add(\"a\"); list.sort(java.util.Comparator.comparing((String s) -> s, java.util.Comparator.nullsLast(String::compareTo))); System.out.println(list.get(0)); System.out.println(list.get(2) == null);",
    );
    assert_eq!(out, vec!["a", "true"]);
}

#[test]
fn lambda_in_stream_filter_with_captured_threshold() {
    let out = run_main(
        "int threshold = 3; long count = java.util.Arrays.asList(1, 2, 3, 4, 5).stream().filter(n -> n > threshold).count(); System.out.println(count);",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn lambda_in_stream_map_with_captured_multiplier() {
    let out = run_main(
        "int factor = 3; java.util.List<Integer> out = java.util.Arrays.asList(1, 2).stream().map(n -> n * factor).toList(); System.out.println(out.get(0)); System.out.println(out.get(1));",
    );
    assert_eq!(out, vec!["3", "6"]);
}

#[test]
fn comparator_chained_then_comparing_three_levels() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"b2\"); list.add(\"a2\"); list.add(\"a10\"); list.sort(java.util.Comparator.comparing((String s) -> s.charAt(0)).thenComparing((String s) -> s.length()).thenComparing(s -> s)); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["a2", "b2"]);
}

#[test]
fn lambda_as_comparator_in_arrays_sort_on_primitive_boxed_array() {
    let out = run_main(
        "Integer[] nums = new Integer[]{3, 1, 2}; java.util.Arrays.sort(nums, (a, b) -> a - b); System.out.println(nums[0]); System.out.println(nums[2]);",
    );
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn lambda_effectively_final_from_ternary_result() {
    let out = run_main(
        "boolean flag = true; int picked = flag ? 8 : 9; java.util.function.Supplier<Integer> read = () -> picked; System.out.println(read.get());",
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn lambda_in_optional_filter_with_captured_bound() {
    let out = run_main(
        "int min = 5; java.util.Optional<Integer> opt = java.util.Optional.of(7).filter(n -> n >= min); System.out.println(opt.get());",
    );
    assert_eq!(out, vec!["7"]);
}
