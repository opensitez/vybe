use crate::helpers::run_main;

#[test]
fn stream_flat_map_splits_strings_into_characters() {
    let out = run_main(
        "long count = java.util.Arrays.asList(\"ab\", \"cd\").stream().flatMap(s -> java.util.Arrays.asList(s.split(\"\")).stream()).filter(c -> c.length() > 0).count(); System.out.println(count);",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn stream_flat_map_on_nested_lists_flattens_one_level() {
    let out = run_main(
        "java.util.List<java.util.List<Integer>> nested = java.util.Arrays.asList(java.util.Arrays.asList(1, 2), java.util.Arrays.asList(3)); long count = nested.stream().flatMap(java.util.Collection::stream).count(); System.out.println(count);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn stream_flat_map_doubles_each_element_via_singleton_stream() {
    let out = run_main(
        "java.util.List<Integer> out = java.util.Arrays.asList(1, 2, 3).stream().flatMap(n -> java.util.Arrays.asList(n, n).stream()).toList(); System.out.println(out.size()); System.out.println(out.get(3));",
    );
    assert_eq!(out, vec!["6", "3"]);
}

#[test]
fn stream_flat_map_on_empty_inner_streams_skips_elements() {
    let out = run_main(
        "long count = java.util.Arrays.asList(1, 2, 3).stream().flatMap(n -> java.util.Arrays.asList(Integer.valueOf(0)).stream().filter(x -> x > n)).count(); System.out.println(count);",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn stream_map_to_int_sum_adds_primitive_values() {
    let out = run_main(
        "int sum = java.util.Arrays.asList(1, 2, 3, 4).stream().mapToInt(n -> n).sum(); System.out.println(sum);",
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn stream_map_to_int_max_finds_largest_mapped_value() {
    let out = run_main(
        "int max = java.util.Arrays.asList(3, 9, 4, 7).stream().mapToInt(n -> n).max().getAsInt(); System.out.println(max);",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn stream_map_to_int_average_computes_mean() {
    let out = run_main(
        "double avg = java.util.Arrays.asList(2, 4, 6).stream().mapToInt(n -> n).average().getAsDouble(); System.out.println((int) avg);",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn stream_map_to_int_after_map_squares_before_sum() {
    let out = run_main(
        "int sum = java.util.Arrays.asList(1, 2, 3).stream().mapToInt(n -> n * n).sum(); System.out.println(sum);",
    );
    assert_eq!(out, vec!["14"]);
}

#[test]
fn stream_distinct_removes_duplicate_integers() {
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
fn stream_distinct_then_sorted_orders_unique_values() {
    let out = run_main(
        "java.util.List<Integer> out = java.util.Arrays.asList(3, 1, 3, 2, 1).stream().distinct().sorted().toList(); System.out.println(out.get(0)); System.out.println(out.get(2));",
    );
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn stream_sorted_orders_integers_ascending() {
    let out = run_main(
        "java.util.List<Integer> sorted = java.util.Arrays.asList(3, 1, 4, 2).stream().sorted().toList(); System.out.println(sorted.get(0)); System.out.println(sorted.get(3));",
    );
    assert_eq!(out, vec!["1", "4"]);
}

#[test]
fn stream_sorted_with_comparator_orders_descending() {
    let out = run_main(
        "java.util.List<Integer> sorted = java.util.Arrays.asList(1, 3, 2).stream().sorted((a, b) -> b - a).toList(); System.out.println(sorted.get(0)); System.out.println(sorted.get(2));",
    );
    assert_eq!(out, vec!["3", "1"]);
}

#[test]
fn stream_sorted_strings_orders_lexicographically() {
    let out = run_main(
        "java.util.List<String> sorted = java.util.Arrays.asList(\"cherry\", \"apple\", \"banana\").stream().sorted().toList(); System.out.println(sorted.get(0)); System.out.println(sorted.get(2));",
    );
    assert_eq!(out, vec!["apple", "cherry"]);
}

#[test]
fn stream_limit_truncates_to_first_n_elements() {
    let out = run_main(
        "java.util.List<Integer> limited = java.util.Arrays.asList(1, 2, 3, 4, 5).stream().limit(3).toList(); System.out.println(limited.size()); System.out.println(limited.get(2));",
    );
    assert_eq!(out, vec!["3", "3"]);
}

#[test]
fn stream_limit_zero_produces_empty_list() {
    let out = run_main(
        "java.util.List<Integer> empty = java.util.Arrays.asList(1, 2, 3).stream().limit(0).toList(); System.out.println(empty.size());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn stream_limit_after_sorted_takes_smallest_values() {
    let out = run_main(
        "java.util.List<Integer> top = java.util.Arrays.asList(5, 1, 4, 2, 3).stream().sorted().limit(2).toList(); System.out.println(top.get(0)); System.out.println(top.get(1));",
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn stream_skip_drops_leading_elements() {
    let out = run_main(
        "java.util.List<Integer> rest = java.util.Arrays.asList(1, 2, 3, 4).stream().skip(2).toList(); System.out.println(rest.get(0)); System.out.println(rest.get(1));",
    );
    assert_eq!(out, vec!["3", "4"]);
}

#[test]
fn stream_skip_all_elements_yields_empty_list() {
    let out = run_main(
        "java.util.List<Integer> empty = java.util.Arrays.asList(1, 2).stream().skip(5).toList(); System.out.println(empty.size());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn stream_skip_then_limit_takes_middle_window() {
    let out = run_main(
        "java.util.List<Integer> window = java.util.Arrays.asList(1, 2, 3, 4, 5).stream().skip(1).limit(2).toList(); System.out.println(window.get(0)); System.out.println(window.get(1));",
    );
    assert_eq!(out, vec!["2", "3"]);
}

#[test]
fn stream_take_while_stops_at_first_non_matching_element() {
    let out = run_main(
        "java.util.List<Integer> taken = java.util.Arrays.asList(1, 2, 3, 1, 2).stream().takeWhile(n -> n < 3).toList(); System.out.println(taken.size()); System.out.println(taken.get(2));",
    );
    assert_eq!(out, vec!["3", "3"]);
}

#[test]
fn stream_take_while_on_all_matching_elements_keeps_entire_stream() {
    let out = run_main(
        "java.util.List<Integer> taken = java.util.Arrays.asList(1, 2, 3).stream().takeWhile(n -> n < 10).toList(); System.out.println(taken.size());",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn stream_take_while_on_empty_prefix_yields_empty_list() {
    let out = run_main(
        "java.util.List<Integer> taken = java.util.Arrays.asList(5, 1, 2).stream().takeWhile(n -> n < 3).toList(); System.out.println(taken.size());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn stream_drop_while_skips_initial_matching_prefix() {
    let out = run_main(
        "java.util.List<Integer> rest = java.util.Arrays.asList(1, 2, 3, 1, 2).stream().dropWhile(n -> n < 3).toList(); System.out.println(rest.size()); System.out.println(rest.get(0));",
    );
    assert_eq!(out, vec!["2", "3"]);
}

#[test]
fn stream_drop_while_when_none_match_keeps_all_elements() {
    let out = run_main(
        "java.util.List<Integer> rest = java.util.Arrays.asList(3, 4, 5).stream().dropWhile(n -> n < 2).toList(); System.out.println(rest.size()); System.out.println(rest.get(2));",
    );
    assert_eq!(out, vec!["3", "5"]);
}

#[test]
fn stream_drop_while_when_all_match_yields_empty_list() {
    let out = run_main(
        "java.util.List<Integer> rest = java.util.Arrays.asList(1, 2, 3).stream().dropWhile(n -> n < 10).toList(); System.out.println(rest.size());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn stream_peek_observes_elements_without_changing_them() {
    let out = run_main(
        "java.util.List<Integer> out = java.util.Arrays.asList(1, 2, 3).stream().peek(n -> System.out.println(n)).toList(); System.out.println(out.get(2));",
    );
    assert_eq!(out, vec!["1", "2", "3", "3"]);
}

#[test]
fn stream_peek_before_filter_only_visits_survivors() {
    let out = run_main(
        "long count = java.util.Arrays.asList(1, 2, 3, 4).stream().peek(n -> System.out.println(n)).filter(n -> n % 2 == 0).count(); System.out.println(count);",
    );
    assert_eq!(out, vec!["1", "2", "3", "4", "2"]);
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
fn stream_reduce_without_identity_returns_optional_sum() {
    let out = run_main(
        "java.util.Optional<Integer> sum = java.util.Arrays.asList(1, 2, 3).stream().reduce((a, b) -> a + b); System.out.println(sum.get());",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn stream_reduce_finds_maximum_value() {
    let out = run_main(
        "int max = java.util.Arrays.asList(3, 9, 4, 7).stream().reduce(0, (a, b) -> a > b ? a : b); System.out.println(max);",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn stream_reduce_with_strings_concatenates_parts() {
    let out = run_main(
        "String joined = java.util.Arrays.asList(\"a\", \"b\", \"c\").stream().reduce(\"\", (a, b) -> a + b); System.out.println(joined);",
    );
    assert_eq!(out, vec!["abc"]);
}

#[test]
fn stream_find_first_returns_first_element_in_encounter_order() {
    let out = run_main(
        "java.util.Optional<Integer> first = java.util.Arrays.asList(5, 1, 4).stream().filter(n -> n > 2).findFirst(); System.out.println(first.get());",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn stream_find_first_on_empty_after_filter_is_not_present() {
    let out = run_main(
        "java.util.Optional<Integer> first = java.util.Arrays.asList(1, 2).stream().filter(n -> n > 5).findFirst(); System.out.println(first.isPresent());",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn stream_find_any_returns_some_matching_element() {
    let out = run_main(
        "java.util.Optional<Integer> any = java.util.Arrays.asList(1, 2, 3, 4).stream().filter(n -> n % 2 == 0).findAny(); System.out.println(any.get());",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn stream_find_any_on_empty_stream_is_not_present() {
    let out = run_main(
        "java.util.Optional<Integer> any = java.util.Arrays.asList(1, 3, 5).stream().filter(n -> n % 2 == 0).findAny(); System.out.println(any.isPresent());",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn stream_any_match_detects_existing_satisfier() {
    let out = run_main(
        "boolean found = java.util.Arrays.asList(1, 3, 4, 5).stream().anyMatch(n -> n % 2 == 0); System.out.println(found);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn stream_any_match_returns_false_when_none_satisfy() {
    let out = run_main(
        "boolean found = java.util.Arrays.asList(1, 3, 5).stream().anyMatch(n -> n % 2 == 0); System.out.println(found);",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn stream_all_match_true_when_every_element_satisfies() {
    let out = run_main(
        "boolean all = java.util.Arrays.asList(2, 4, 6).stream().allMatch(n -> n % 2 == 0); System.out.println(all);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn stream_all_match_false_when_one_element_fails() {
    let out = run_main(
        "boolean all = java.util.Arrays.asList(2, 3, 4).stream().allMatch(n -> n % 2 == 0); System.out.println(all);",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn stream_none_match_true_when_no_element_satisfies() {
    let out = run_main(
        "boolean none = java.util.Arrays.asList(1, 3, 5).stream().noneMatch(n -> n % 2 == 0); System.out.println(none);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn stream_none_match_false_when_at_least_one_satisfies() {
    let out = run_main(
        "boolean none = java.util.Arrays.asList(1, 2, 3).stream().noneMatch(n -> n % 2 == 0); System.out.println(none);",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn stream_flat_map_then_map_transforms_flattened_elements() {
    let out = run_main(
        "java.util.List<Integer> out = java.util.Arrays.asList(1, 2).stream().flatMap(n -> java.util.Arrays.asList(n, n + 10).stream()).map(n -> n * 2).toList(); System.out.println(out.get(1)); System.out.println(out.get(3));",
    );
    assert_eq!(out, vec!["4", "24"]);
}

#[test]
fn stream_match_combinations_short_circuit_on_first_failure() {
    let out = run_main(
        "boolean allPositive = java.util.Arrays.asList(1, 2, -3).stream().allMatch(n -> n > 0); boolean anyPositive = java.util.Arrays.asList(-1, -2, 3).stream().anyMatch(n -> n > 0); System.out.println(allPositive); System.out.println(anyPositive);",
    );
    assert_eq!(out, vec!["false", "true"]);
}
