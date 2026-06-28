use crate::helpers::run_main;

#[test]
fn stream_map_doubles_each_integer() {
    let out = run_main(
        "java.util.List<Integer> nums = java.util.Arrays.asList(1, 2, 3); java.util.List<Integer> doubled = nums.stream().map(n -> n * 2).toList(); System.out.println(doubled.get(0)); System.out.println(doubled.get(2));",
    );
    assert_eq!(out, vec!["2", "6"]);
}

#[test]
fn stream_map_converts_integers_to_strings() {
    let out = run_main(
        "java.util.List<String> labels = java.util.Arrays.asList(1, 2, 3).stream().map(n -> \"v\" + n).toList(); System.out.println(labels.get(1));",
    );
    assert_eq!(out, vec!["v2"]);
}

#[test]
fn stream_filter_keeps_even_numbers_only() {
    let out = run_main(
        "long count = java.util.Arrays.asList(1, 2, 3, 4).stream().filter(n -> n % 2 == 0).count(); System.out.println(count);",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn stream_filter_on_empty_list_yields_zero_count() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); long count = list.stream().filter(n -> n > 0).count(); System.out.println(count);",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn stream_count_returns_total_elements_without_filter() {
    let out = run_main(
        "long count = java.util.Arrays.asList(5, 6, 7, 8, 9).stream().count(); System.out.println(count);",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn stream_reduce_with_identity_sums_integers() {
    let out = run_main(
        "int sum = java.util.Arrays.asList(1, 2, 3, 4).stream().reduce(0, (a, b) -> a + b); System.out.println(sum);",
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn stream_reduce_multiplies_accumulator() {
    let out = run_main(
        "int product = java.util.Arrays.asList(2, 3, 4).stream().reduce(1, (a, b) -> a * b); System.out.println(product);",
    );
    assert_eq!(out, vec!["24"]);
}

#[test]
fn stream_reduce_on_single_element_returns_that_element() {
    let out = run_main(
        "int only = java.util.Arrays.asList(99).stream().reduce(0, (a, b) -> a + b); System.out.println(only);",
    );
    assert_eq!(out, vec!["99"]);
}

#[test]
fn stream_collect_to_list_preserves_order() {
    let out = run_main(
        "java.util.List<Integer> out = java.util.Arrays.asList(3, 1, 2).stream().collect(java.util.stream.Collectors.toList()); System.out.println(out.get(0)); System.out.println(out.get(2));",
    );
    assert_eq!(out, vec!["3", "2"]);
}

#[test]
fn stream_to_list_materializes_transformed_values() {
    let out = run_main(
        "java.util.List<Integer> tripled = java.util.Arrays.asList(1, 2).stream().map(n -> n * 3).toList(); System.out.println(tripled.size()); System.out.println(tripled.get(1));",
    );
    assert_eq!(out, vec!["2", "6"]);
}

#[test]
fn stream_sorted_orders_integers_ascending() {
    let out = run_main(
        "java.util.List<Integer> sorted = java.util.Arrays.asList(3, 1, 4, 2).stream().sorted().toList(); System.out.println(sorted.get(0)); System.out.println(sorted.get(3));",
    );
    assert_eq!(out, vec!["1", "4"]);
}

#[test]
fn stream_sorted_after_map_still_orders_results() {
    let out = run_main(
        "java.util.List<Integer> sorted = java.util.Arrays.asList(3, 1, 2).stream().map(n -> n * 10).sorted().toList(); System.out.println(sorted.get(0)); System.out.println(sorted.get(2));",
    );
    assert_eq!(out, vec!["10", "30"]);
}

#[test]
fn stream_distinct_removes_duplicate_values() {
    let out = run_main(
        "long count = java.util.Arrays.asList(1, 2, 2, 3, 3, 3).stream().distinct().count(); System.out.println(count);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn stream_distinct_on_already_unique_list_keeps_size() {
    let out = run_main(
        "long count = java.util.Arrays.asList(4, 5, 6).stream().distinct().count(); System.out.println(count);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn stream_limit_truncates_to_first_n_elements() {
    let out = run_main(
        "java.util.List<Integer> limited = java.util.Arrays.asList(1, 2, 3, 4, 5).stream().limit(3).toList(); System.out.println(limited.size()); System.out.println(limited.get(2));",
    );
    assert_eq!(out, vec!["3", "3"]);
}

#[test]
fn stream_limit_one_returns_singleton_list() {
    let out = run_main(
        "java.util.List<Integer> one = java.util.Arrays.asList(7, 8, 9).stream().limit(1).toList(); System.out.println(one.get(0));",
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn stream_skip_drops_leading_elements() {
    let out = run_main(
        "java.util.List<Integer> rest = java.util.Arrays.asList(1, 2, 3, 4).stream().skip(2).toList(); System.out.println(rest.get(0)); System.out.println(rest.get(1));",
    );
    assert_eq!(out, vec!["3", "4"]);
}

#[test]
fn stream_skip_then_limit_takes_middle_window() {
    let out = run_main(
        "java.util.List<Integer> window = java.util.Arrays.asList(1, 2, 3, 4, 5).stream().skip(1).limit(2).toList(); System.out.println(window.get(0)); System.out.println(window.get(1));",
    );
    assert_eq!(out, vec!["2", "3"]);
}

#[test]
fn stream_foreach_prints_each_element_in_order() {
    let out = run_main(
        "java.util.Arrays.asList(10, 20, 30).stream().forEach(n -> System.out.println(n));",
    );
    assert_eq!(out, vec!["10", "20", "30"]);
}

#[test]
fn stream_foreach_on_filtered_stream_prints_matches_only() {
    let out = run_main(
        "java.util.Arrays.asList(1, 2, 3, 4).stream().filter(n -> n > 2).forEach(n -> System.out.println(n));",
    );
    assert_eq!(out, vec!["3", "4"]);
}

#[test]
fn stream_map_then_filter_composes_transform_and_predicate() {
    let out = run_main(
        "long count = java.util.Arrays.asList(1, 2, 3, 4).stream().map(n -> n * 2).filter(n -> n > 4).count(); System.out.println(count);",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn stream_filter_then_map_keeps_order_of_survivors() {
    let out = run_main(
        "java.util.List<Integer> out = java.util.Arrays.asList(1, 2, 3, 4).stream().filter(n -> n % 2 == 1).map(n -> n + 10).toList(); System.out.println(out.get(0)); System.out.println(out.get(1));",
    );
    assert_eq!(out, vec!["11", "13"]);
}

#[test]
fn stream_from_arraylist_supports_map() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(2); list.add(4); java.util.List<Integer> out = list.stream().map(n -> n + 1).toList(); System.out.println(out.get(1));",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn stream_count_after_distinct_counts_unique_items() {
    let out = run_main(
        "long count = java.util.Arrays.asList(1, 1, 2, 2, 3).stream().distinct().count(); System.out.println(count);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn stream_sorted_strings_orders_lexicographically() {
    let out = run_main(
        "java.util.List<String> sorted = java.util.Arrays.asList(\"cherry\", \"apple\", \"banana\").stream().sorted().toList(); System.out.println(sorted.get(0)); System.out.println(sorted.get(2));",
    );
    assert_eq!(out, vec!["apple", "cherry"]);
}

#[test]
fn stream_map_to_length_on_strings() {
    let out = run_main(
        "java.util.List<Integer> lengths = java.util.Arrays.asList(\"a\", \"ab\", \"abc\").stream().map(s -> s.length()).toList(); System.out.println(lengths.get(2));",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn stream_reduce_finds_maximum_value() {
    let out = run_main(
        "int max = java.util.Arrays.asList(3, 9, 4, 7).stream().reduce(0, (a, b) -> a > b ? a : b); System.out.println(max);",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn stream_filter_negative_numbers_returns_empty_count() {
    let out = run_main(
        "long count = java.util.Arrays.asList(-1, -2, -3).stream().filter(n -> n > 0).count(); System.out.println(count);",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn stream_limit_zero_produces_empty_list() {
    let out = run_main(
        "java.util.List<Integer> empty = java.util.Arrays.asList(1, 2, 3).stream().limit(0).toList(); System.out.println(empty.size());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn stream_skip_all_elements_yields_empty_list() {
    let out = run_main(
        "java.util.List<Integer> empty = java.util.Arrays.asList(1, 2).stream().skip(5).toList(); System.out.println(empty.size());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn stream_distinct_then_sorted_orders_unique_values() {
    let out = run_main(
        "java.util.List<Integer> out = java.util.Arrays.asList(3, 1, 3, 2, 1).stream().distinct().sorted().toList(); System.out.println(out.get(0)); System.out.println(out.get(2));",
    );
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn stream_map_identity_preserves_values() {
    let out = run_main(
        "java.util.List<Integer> same = java.util.Arrays.asList(4, 5, 6).stream().map(n -> n).toList(); System.out.println(same.get(1));",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn stream_foreach_on_sorted_stream_prints_in_order() {
    let out = run_main(
        "java.util.Arrays.asList(3, 1, 2).stream().sorted().forEach(n -> System.out.println(n));",
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn stream_collect_after_filter_builds_sublist() {
    let out = run_main(
        "java.util.List<Integer> evens = java.util.Arrays.asList(1, 2, 3, 4, 5, 6).stream().filter(n -> n % 2 == 0).collect(java.util.stream.Collectors.toList()); System.out.println(evens.size()); System.out.println(evens.get(1));",
    );
    assert_eq!(out, vec!["3", "6"]);
}

#[test]
fn stream_reduce_with_strings_concatenates_parts() {
    let out = run_main(
        "String joined = java.util.Arrays.asList(\"a\", \"b\", \"c\").stream().reduce(\"\", (a, b) -> a + b); System.out.println(joined);",
    );
    assert_eq!(out, vec!["abc"]);
}

#[test]
fn stream_from_list_of_three_counts_three() {
    let out = run_main(
        "java.util.List<Integer> list = java.util.Arrays.asList(9, 8, 7); System.out.println(list.stream().count());",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn stream_filter_greater_than_threshold() {
    let out = run_main(
        "long count = java.util.Arrays.asList(1, 5, 9, 2, 8).stream().filter(n -> n >= 5).count(); System.out.println(count);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn stream_map_adds_constant_offset() {
    let out = run_main(
        "java.util.List<Integer> shifted = java.util.Arrays.asList(1, 2, 3).stream().map(n -> n + 100).toList(); System.out.println(shifted.get(0)); System.out.println(shifted.get(2));",
    );
    assert_eq!(out, vec!["101", "103"]);
}

#[test]
fn stream_skip_one_then_map_doubles_remainder() {
    let out = run_main(
        "java.util.List<Integer> out = java.util.Arrays.asList(1, 2, 3, 4).stream().skip(1).map(n -> n * 2).toList(); System.out.println(out.get(0)); System.out.println(out.get(2));",
    );
    assert_eq!(out, vec!["4", "8"]);
}

#[test]
fn stream_limit_after_sorted_takes_smallest_values() {
    let out = run_main(
        "java.util.List<Integer> top = java.util.Arrays.asList(5, 1, 4, 2, 3).stream().sorted().limit(2).toList(); System.out.println(top.get(0)); System.out.println(top.get(1));",
    );
    assert_eq!(out, vec!["1", "2"]);
}
