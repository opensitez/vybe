use crate::helpers::run_main;

#[test]
fn int_stream_range_excludes_upper_bound() {
    let out = run_main(
        "long c = java.util.stream.IntStream.range(1, 4).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn int_stream_range_first_value() {
    let out = run_main(
        "int v = java.util.stream.IntStream.range(5, 10).findFirst().getAsInt(); System.out.println(v);",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn int_stream_range_last_value_before_end() {
    let out = run_main(
        "int v = java.util.stream.IntStream.range(2, 6).max().getAsInt(); System.out.println(v);",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn int_stream_range_empty_when_start_equals_end() {
    let out = run_main(
        "long c = java.util.stream.IntStream.range(7, 7).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn int_stream_range_closed_includes_upper_bound() {
    let out = run_main(
        "long c = java.util.stream.IntStream.rangeClosed(1, 4).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn int_stream_range_closed_last_element() {
    let out = run_main(
        "int v = java.util.stream.IntStream.rangeClosed(3, 5).max().getAsInt(); System.out.println(v);",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn int_stream_of_three_values_sums() {
    let out = run_main(
        "int s = java.util.stream.IntStream.of(1, 2, 3).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn int_stream_of_single_value_count() {
    let out = run_main(
        "long c = java.util.stream.IntStream.of(42).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn int_stream_empty_yields_zero_count() {
    let out = run_main(
        "long c = java.util.stream.IntStream.empty().count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn int_stream_iterate_doubles_each_step() {
    let out = run_main(
        "int v = java.util.stream.IntStream.iterate(1, n -> n * 2).limit(4).sum(); System.out.println(v);",
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn int_stream_iterate_with_predicate_stops_early() {
    let out = run_main(
        "long c = java.util.stream.IntStream.iterate(1, n -> n < 10, n -> n + 3).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn int_stream_generate_increments_supplier() {
    let out = run_main(
        "int[] box = {0}; int s = java.util.stream.IntStream.generate(() -> ++box[0]).limit(3).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn int_stream_filter_keeps_even_values() {
    let out = run_main(
        "int s = java.util.stream.IntStream.range(1, 7).filter(n -> n % 2 == 0).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn int_stream_filter_discards_all_when_none_match() {
    let out = run_main(
        "long c = java.util.stream.IntStream.range(1, 5).filter(n -> n > 10).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn int_stream_map_multiplies_by_three() {
    let out = run_main(
        "int s = java.util.stream.IntStream.of(2, 3, 4).map(n -> n * 3).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["27"]);
}

#[test]
fn int_stream_map_to_long_preserves_sum() {
    let out = run_main(
        "long s = java.util.stream.IntStream.of(1, 2, 3).mapToLong(n -> (long) n).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn int_stream_map_to_double_averages_halves() {
    let out = run_main(
        "double a = java.util.stream.IntStream.of(2, 4, 6).mapToDouble(n -> n / 2.0).average().getAsDouble(); System.out.println((int) a);",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn int_stream_map_to_obj_boxes_and_counts() {
    let out = run_main(
        "long c = java.util.stream.IntStream.of(1, 2).mapToObj(n -> \"v\" + n).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn int_stream_flat_map_duplicates_each_value() {
    let out = run_main(
        "int s = java.util.stream.IntStream.of(1, 2).flatMap(n -> java.util.stream.IntStream.of(n, n)).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn int_stream_distinct_on_of_with_duplicates() {
    let out = run_main(
        "long c = java.util.stream.IntStream.of(1, 1, 2, 2, 3).distinct().count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn int_stream_sorted_descending_via_boxed() {
    let out = run_main(
        "int first = java.util.stream.IntStream.of(3, 1, 2).boxed().sorted((a, b) -> b - a).findFirst().get(); System.out.println(first);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn int_stream_peek_increments_side_counter() {
    let out = run_main(
        "int[] seen = {0}; int s = java.util.stream.IntStream.of(1, 2, 3).peek(n -> seen[0]++).sum(); System.out.println(seen[0]); System.out.println(s);",
    );
    assert_eq!(out, vec!["3", "6"]);
}

#[test]
fn int_stream_limit_truncates_infinite_iterate() {
    let out = run_main(
        "long c = java.util.stream.IntStream.iterate(0, n -> n + 1).limit(5).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn int_stream_skip_drops_leading_elements() {
    let out = run_main(
        "int s = java.util.stream.IntStream.range(1, 6).skip(2).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn int_stream_for_each_prints_sequence() {
    let out = run_main(
        "java.util.stream.IntStream.of(4, 5).forEach(n -> System.out.println(n));",
    );
    assert_eq!(out, vec!["4", "5"]);
}

#[test]
fn int_stream_to_array_length_matches_count() {
    let out = run_main(
        "int len = java.util.stream.IntStream.range(1, 4).toArray().length; System.out.println(len);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn int_stream_to_array_first_element() {
    let out = run_main(
        "int v = java.util.stream.IntStream.of(9, 8, 7).toArray()[0]; System.out.println(v);",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn int_stream_reduce_with_identity_adds_all() {
    let out = run_main(
        "int s = java.util.stream.IntStream.of(1, 2, 3, 4).reduce(0, (a, b) -> a + b); System.out.println(s);",
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn int_stream_reduce_without_identity_multiplies() {
    let out = run_main(
        "int p = java.util.stream.IntStream.of(2, 3, 4).reduce((a, b) -> a * b).getAsInt(); System.out.println(p);",
    );
    assert_eq!(out, vec!["24"]);
}

#[test]
fn int_stream_min_finds_smallest_in_range() {
    let out = run_main(
        "int v = java.util.stream.IntStream.of(7, 2, 9, 1).min().getAsInt(); System.out.println(v);",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn int_stream_max_finds_largest_in_range() {
    let out = run_main(
        "int v = java.util.stream.IntStream.of(7, 2, 9, 1).max().getAsInt(); System.out.println(v);",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn int_stream_average_of_consecutive_triplet() {
    let out = run_main(
        "double a = java.util.stream.IntStream.of(2, 5, 8).average().getAsDouble(); System.out.println((int) a);",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn int_stream_sum_of_squares() {
    let out = run_main(
        "int s = java.util.stream.IntStream.of(1, 2, 3).map(n -> n * n).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["14"]);
}

#[test]
fn int_stream_count_after_filter() {
    let out = run_main(
        "long c = java.util.stream.IntStream.range(1, 11).filter(n -> n % 3 == 0).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn int_stream_any_match_detects_greater_than_five() {
    let out = run_main(
        "boolean b = java.util.stream.IntStream.of(1, 2, 6).anyMatch(n -> n > 5); System.out.println(b);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn int_stream_any_match_false_on_empty_stream() {
    let out = run_main(
        "boolean b = java.util.stream.IntStream.empty().anyMatch(n -> n > 0); System.out.println(b);",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn int_stream_all_match_true_for_all_positive() {
    let out = run_main(
        "boolean b = java.util.stream.IntStream.of(1, 2, 3).allMatch(n -> n > 0); System.out.println(b);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn int_stream_all_match_false_when_one_non_positive() {
    let out = run_main(
        "boolean b = java.util.stream.IntStream.of(1, 0, 3).allMatch(n -> n > 0); System.out.println(b);",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn int_stream_none_match_true_when_all_below_zero() {
    let out = run_main(
        "boolean b = java.util.stream.IntStream.of(-1, -2).noneMatch(n -> n > 0); System.out.println(b);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn int_stream_none_match_false_when_positive_exists() {
    let out = run_main(
        "boolean b = java.util.stream.IntStream.of(-1, 2).noneMatch(n -> n > 0); System.out.println(b);",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn int_stream_find_first_returns_first_match() {
    let out = run_main(
        "int v = java.util.stream.IntStream.of(1, 3, 5, 6).filter(n -> n % 2 == 0).findFirst().getAsInt(); System.out.println(v);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn int_stream_find_any_returns_some_match() {
    let out = run_main(
        "int v = java.util.stream.IntStream.of(9, 8, 7).filter(n -> n < 9).findAny().getAsInt(); System.out.println(v);",
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn int_stream_boxed_collects_to_list_size() {
    let out = run_main(
        "int sz = java.util.stream.IntStream.of(1, 2, 3).boxed().toList().size(); System.out.println(sz);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn int_stream_as_long_stream_sums() {
    let out = run_main(
        "long s = java.util.stream.IntStream.of(1, 2, 3).asLongStream().sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn int_stream_as_double_stream_averages() {
    let out = run_main(
        "double a = java.util.stream.IntStream.of(2, 4, 6).asDoubleStream().average().getAsDouble(); System.out.println((int) a);",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn int_stream_concat_two_ranges() {
    let out = run_main(
        "long c = java.util.stream.IntStream.concat(java.util.stream.IntStream.of(1, 2), java.util.stream.IntStream.of(3)).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn int_stream_builder_accumulates_elements() {
    let out = run_main(
        "int s = java.util.stream.IntStream.builder().add(2).add(3).build().sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn int_stream_summary_statistics_sum() {
    let out = run_main(
        "int s = (int) java.util.stream.IntStream.of(1, 2, 3).summaryStatistics().getSum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn int_stream_summary_statistics_max() {
    let out = run_main(
        "int v = java.util.stream.IntStream.of(1, 5, 3).summaryStatistics().getMax(); System.out.println(v);",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn int_stream_summary_statistics_average() {
    let out = run_main(
        "double a = java.util.stream.IntStream.of(2, 4, 6).summaryStatistics().getAverage(); System.out.println((int) a);",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn int_stream_summary_statistics_count() {
    let out = run_main(
        "long c = java.util.stream.IntStream.of(1, 2, 3, 4).summaryStatistics().getCount(); System.out.println(c);",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn int_stream_collect_joins_string_representations() {
    let out = run_main(
        "String j = java.util.stream.IntStream.of(1, 2, 3).mapToObj(n -> \"\" + n).collect(java.util.stream.Collectors.joining(\"\")); System.out.println(j);",
    );
    assert_eq!(out, vec!["123"]);
}

#[test]
fn int_stream_drop_while_skips_prefix() {
    let out = run_main(
        "int s = java.util.stream.IntStream.of(1, 2, 3, 1, 2).dropWhile(n -> n < 3).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn int_stream_take_while_keeps_prefix() {
    let out = run_main(
        "int s = java.util.stream.IntStream.of(1, 2, 3, 4).takeWhile(n -> n < 3).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn int_stream_reduce_builds_decimal_number() {
    let out = run_main(
        "int built = java.util.stream.IntStream.of(1, 2, 3).reduce(0, (a, b) -> a * 10 + b); System.out.println(built);",
    );
    assert_eq!(out, vec!["123"]);
}

#[test]
fn int_stream_iterate_fibonacci_third_value() {
    let out = run_main(
        "int[] prev = {0, 1}; int v = java.util.stream.IntStream.iterate(0, n -> { int next = prev[0] + prev[1]; prev[0] = prev[1]; prev[1] = next; return next; }).limit(4).reduce((a, b) -> b).getAsInt(); System.out.println(v);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn int_stream_range_sum_one_to_five() {
    let out = run_main(
        "int s = java.util.stream.IntStream.rangeClosed(1, 5).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn int_stream_map_abs_on_negative() {
    let out = run_main(
        "int s = java.util.stream.IntStream.of(-2, -1, 3).map(Math::abs).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn int_stream_filter_then_map_chain() {
    let out = run_main(
        "int s = java.util.stream.IntStream.range(1, 6).filter(n -> n % 2 == 1).map(n -> n * 10).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["90"]);
}

#[test]
fn int_stream_distinct_on_constant_generate() {
    let out = run_main(
        "long c = java.util.stream.IntStream.generate(() -> 7).limit(4).distinct().count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn int_stream_find_first_empty_when_no_match() {
    let out = run_main(
        "boolean present = java.util.stream.IntStream.of(1, 3, 5).filter(n -> n % 2 == 0).findFirst().isPresent(); System.out.println(present);",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn int_stream_reduce_optional_on_singleton() {
    let out = run_main(
        "int v = java.util.stream.IntStream.of(99).reduce((a, b) -> a + b).getAsInt(); System.out.println(v);",
    );
    assert_eq!(out, vec!["99"]);
}

#[test]
fn int_stream_reduce_optional_empty_returns_empty() {
    let out = run_main(
        "boolean present = java.util.stream.IntStream.empty().reduce((a, b) -> a + b).isPresent(); System.out.println(present);",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn int_stream_skip_then_limit_window() {
    let out = run_main(
        "int s = java.util.stream.IntStream.range(1, 10).skip(3).limit(2).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn int_stream_peek_does_not_change_sum() {
    let out = run_main(
        "int s = java.util.stream.IntStream.of(2, 3).peek(n -> n + 100).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn int_stream_to_array_last_element() {
    let out = run_main(
        "int[] arr = java.util.stream.IntStream.of(4, 5, 6).toArray(); System.out.println(arr[arr.length - 1]);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn int_stream_generate_constant_limit_three() {
    let out = run_main(
        "int s = java.util.stream.IntStream.generate(() -> 5).limit(3).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn int_stream_range_negative_to_positive() {
    let out = run_main(
        "int s = java.util.stream.IntStream.rangeClosed(-2, 2).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn int_stream_of_spread_max_value() {
    let out = run_main(
        "int v = java.util.stream.IntStream.of(3, 9, 5).max().getAsInt(); System.out.println(v);",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn int_stream_match_short_circuits_on_first_true() {
    let out = run_main(
        "boolean b = java.util.stream.IntStream.of(1, 2, 3).anyMatch(n -> n == 2); System.out.println(b);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn int_stream_sequential_flag_still_sums() {
    let out = run_main(
        "int s = java.util.stream.IntStream.of(1, 2, 3).sequential().sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn int_stream_map_to_long_then_back_sum() {
    let out = run_main(
        "long s = java.util.stream.IntStream.of(1, 2).mapToLong(n -> n * 10L).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["30"]);
}

#[test]
fn int_stream_flat_map_to_empty_skips() {
    let out = run_main(
        "long c = java.util.stream.IntStream.of(1, 2, 3).flatMap(n -> java.util.stream.IntStream.empty()).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn int_stream_filter_map_sum_pipeline() {
    let out = run_main(
        "int s = java.util.stream.IntStream.rangeClosed(1, 10).filter(n -> n <= 4).map(n -> n + 1).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["14"]);
}

