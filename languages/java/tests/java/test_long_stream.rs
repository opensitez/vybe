use crate::helpers::run_main;

#[test]
fn long_stream_range_excludes_upper_bound() {
    let out = run_main(
        "long c = java.util.stream.LongStream.range(1L, 4L).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn long_stream_range_first_value() {
    let out = run_main(
        "long v = java.util.stream.LongStream.range(5L, 10L).findFirst().getAsLong(); System.out.println(v);",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn long_stream_range_closed_includes_end() {
    let out = run_main(
        "long c = java.util.stream.LongStream.rangeClosed(1L, 4L).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn long_stream_range_closed_last() {
    let out = run_main(
        "long v = java.util.stream.LongStream.rangeClosed(3L, 5L).max().getAsLong(); System.out.println(v);",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn long_stream_of_three_sums() {
    let out = run_main(
        "long s = java.util.stream.LongStream.of(1L, 2L, 3L).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn long_stream_of_single_count() {
    let out =
        run_main("long c = java.util.stream.LongStream.of(42L).count(); System.out.println(c);");
    assert_eq!(out, vec!["1"]);
}

#[test]
fn long_stream_empty_count_zero() {
    let out =
        run_main("long c = java.util.stream.LongStream.empty().count(); System.out.println(c);");
    assert_eq!(out, vec!["0"]);
}

#[test]
fn long_stream_iterate_doubles() {
    let out = run_main(
        "long s = java.util.stream.LongStream.iterate(1L, n -> n * 2L).limit(4L).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn long_stream_iterate_with_predicate() {
    let out = run_main(
        "long c = java.util.stream.LongStream.iterate(1L, n -> n < 10L, n -> n + 3L).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn long_stream_generate_increments() {
    let out = run_main(
        "long[] box = {0L}; long s = java.util.stream.LongStream.generate(() -> ++box[0]).limit(3L).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn long_stream_filter_evens() {
    let out = run_main(
        "long s = java.util.stream.LongStream.range(1L, 7L).filter(n -> n % 2L == 0L).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn long_stream_filter_none_match() {
    let out = run_main(
        "long c = java.util.stream.LongStream.range(1L, 5L).filter(n -> n > 10L).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn long_stream_map_multiplies() {
    let out = run_main(
        "long s = java.util.stream.LongStream.of(2L, 3L, 4L).map(n -> n * 3L).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["27"]);
}

#[test]
fn long_stream_map_to_int_truncates_sum() {
    let out = run_main(
        "int s = java.util.stream.LongStream.of(2L, 3L, 4L).mapToInt(n -> (int) n).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn long_stream_map_to_double_halves() {
    let out = run_main(
        "double a = java.util.stream.LongStream.of(4L, 8L).mapToDouble(n -> n / 2.0).average().getAsDouble(); System.out.println((int) a);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn long_stream_map_to_obj_count() {
    let out = run_main(
        "long c = java.util.stream.LongStream.of(1L, 2L).mapToObj(n -> \"L\" + n).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn long_stream_flat_map_duplicates() {
    let out = run_main(
        "long s = java.util.stream.LongStream.of(1L, 2L).flatMap(n -> java.util.stream.LongStream.of(n, n)).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn long_stream_distinct_on_duplicates() {
    let out = run_main(
        "long c = java.util.stream.LongStream.of(1L, 1L, 2L, 2L, 3L).distinct().count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn long_stream_sorted_desc_via_boxed() {
    let out = run_main(
        "long first = java.util.stream.LongStream.of(3L, 1L, 2L).boxed().sorted((a, b) -> Long.compare(b, a)).findFirst().get(); System.out.println(first);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn long_stream_peek_counts_elements() {
    let out = run_main(
        "long[] seen = {0L}; long s = java.util.stream.LongStream.of(1L, 2L, 3L).peek(n -> seen[0]++).sum(); System.out.println(seen[0]); System.out.println(s);",
    );
    assert_eq!(out, vec!["3", "6"]);
}

#[test]
fn long_stream_limit_on_iterate() {
    let out = run_main(
        "long c = java.util.stream.LongStream.iterate(0L, n -> n + 1L).limit(5L).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn long_stream_skip_drops_prefix() {
    let out = run_main(
        "long s = java.util.stream.LongStream.range(1L, 6L).skip(2L).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn long_stream_for_each_prints() {
    let out =
        run_main("java.util.stream.LongStream.of(4L, 5L).forEach(n -> System.out.println(n));");
    assert_eq!(out, vec!["4", "5"]);
}

#[test]
fn long_stream_to_array_length() {
    let out = run_main(
        "int len = java.util.stream.LongStream.range(1L, 4L).toArray().length; System.out.println(len);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn long_stream_reduce_with_identity() {
    let out = run_main(
        "long s = java.util.stream.LongStream.of(1L, 2L, 3L, 4L).reduce(0L, (a, b) -> a + b); System.out.println(s);",
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn long_stream_reduce_without_identity() {
    let out = run_main(
        "long p = java.util.stream.LongStream.of(2L, 3L, 4L).reduce((a, b) -> a * b).getAsLong(); System.out.println(p);",
    );
    assert_eq!(out, vec!["24"]);
}

#[test]
fn long_stream_min_value() {
    let out = run_main(
        "long v = java.util.stream.LongStream.of(7L, 2L, 9L, 1L).min().getAsLong(); System.out.println(v);",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn long_stream_max_value() {
    let out = run_main(
        "long v = java.util.stream.LongStream.of(7L, 2L, 9L, 1L).max().getAsLong(); System.out.println(v);",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn long_stream_average_triplet() {
    let out = run_main(
        "double a = java.util.stream.LongStream.of(2L, 5L, 8L).average().getAsDouble(); System.out.println((int) a);",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn long_stream_sum_of_squares() {
    let out = run_main(
        "long s = java.util.stream.LongStream.of(1L, 2L, 3L).map(n -> n * n).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["14"]);
}

#[test]
fn long_stream_count_multiples_of_three() {
    let out = run_main(
        "long c = java.util.stream.LongStream.range(1L, 11L).filter(n -> n % 3L == 0L).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn long_stream_any_match_true() {
    let out = run_main(
        "boolean b = java.util.stream.LongStream.of(1L, 2L, 6L).anyMatch(n -> n > 5L); System.out.println(b);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn long_stream_any_match_false_empty() {
    let out = run_main(
        "boolean b = java.util.stream.LongStream.empty().anyMatch(n -> n > 0L); System.out.println(b);",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn long_stream_all_match_true() {
    let out = run_main(
        "boolean b = java.util.stream.LongStream.of(1L, 2L, 3L).allMatch(n -> n > 0L); System.out.println(b);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn long_stream_all_match_false() {
    let out = run_main(
        "boolean b = java.util.stream.LongStream.of(1L, 0L, 3L).allMatch(n -> n > 0L); System.out.println(b);",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn long_stream_none_match_true() {
    let out = run_main(
        "boolean b = java.util.stream.LongStream.of(-1L, -2L).noneMatch(n -> n > 0L); System.out.println(b);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn long_stream_none_match_false() {
    let out = run_main(
        "boolean b = java.util.stream.LongStream.of(-1L, 2L).noneMatch(n -> n > 0L); System.out.println(b);",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn long_stream_find_first_even() {
    let out = run_main(
        "long v = java.util.stream.LongStream.of(1L, 3L, 5L, 6L).filter(n -> n % 2L == 0L).findFirst().getAsLong(); System.out.println(v);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn long_stream_find_any_less_than_nine() {
    let out = run_main(
        "long v = java.util.stream.LongStream.of(9L, 8L, 7L).filter(n -> n < 9L).findAny().getAsLong(); System.out.println(v);",
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn long_stream_boxed_list_size() {
    let out = run_main(
        "int sz = java.util.stream.LongStream.of(1L, 2L, 3L).boxed().toList().size(); System.out.println(sz);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn long_stream_as_int_stream_sums() {
    let out = run_main(
        "int s = java.util.stream.LongStream.of(1L, 2L, 3L).mapToInt(n -> (int) n).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn long_stream_as_double_stream_avg() {
    let out = run_main(
        "double a = java.util.stream.LongStream.of(2L, 4L, 6L).asDoubleStream().average().getAsDouble(); System.out.println((int) a);",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn long_stream_concat_two() {
    let out = run_main(
        "long c = java.util.stream.LongStream.concat(java.util.stream.LongStream.of(1L, 2L), java.util.stream.LongStream.of(3L)).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn long_stream_builder_sum() {
    let out = run_main(
        "long s = java.util.stream.LongStream.builder().add(2L).add(3L).build().sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn long_stream_summary_statistics_sum() {
    let out = run_main(
        "long s = java.util.stream.LongStream.of(1L, 2L, 3L).summaryStatistics().getSum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn long_stream_summary_statistics_max() {
    let out = run_main(
        "long v = java.util.stream.LongStream.of(1L, 5L, 3L).summaryStatistics().getMax(); System.out.println(v);",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn long_stream_summary_statistics_avg() {
    let out = run_main(
        "double a = java.util.stream.LongStream.of(2L, 4L, 6L).summaryStatistics().getAverage(); System.out.println((int) a);",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn long_stream_summary_statistics_count() {
    let out = run_main(
        "long c = java.util.stream.LongStream.of(1L, 2L, 3L, 4L).summaryStatistics().getCount(); System.out.println(c);",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn long_stream_collect_joining() {
    let out = run_main(
        "String j = java.util.stream.LongStream.of(1L, 2L, 3L).mapToObj(n -> \"\" + n).collect(java.util.stream.Collectors.joining(\"\")); System.out.println(j);",
    );
    assert_eq!(out, vec!["123"]);
}

#[test]
fn long_stream_drop_while_prefix() {
    let out = run_main(
        "long s = java.util.stream.LongStream.of(1L, 2L, 3L, 1L).dropWhile(n -> n < 3L).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn long_stream_take_while_prefix() {
    let out = run_main(
        "long s = java.util.stream.LongStream.of(1L, 2L, 3L, 4L).takeWhile(n -> n < 3L).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn long_stream_reduce_builds_decimal_number() {
    let out = run_main(
        "long built = java.util.stream.LongStream.of(1L, 2L, 3L).reduce(0L, (a, b) -> a * 10L + b); System.out.println(built);",
    );
    assert_eq!(out, vec!["123"]);
}

#[test]
fn long_stream_range_closed_sum_one_to_five() {
    let out = run_main(
        "long s = java.util.stream.LongStream.rangeClosed(1L, 5L).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn long_stream_map_abs_negative() {
    let out = run_main(
        "long s = java.util.stream.LongStream.of(-2L, -1L, 3L).map(Math::abs).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn long_stream_filter_map_chain() {
    let out = run_main(
        "long s = java.util.stream.LongStream.range(1L, 6L).filter(n -> n % 2L == 1L).map(n -> n * 10L).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["90"]);
}

#[test]
fn long_stream_distinct_constant_generate() {
    let out = run_main(
        "long c = java.util.stream.LongStream.generate(() -> 7L).limit(4L).distinct().count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn long_stream_find_first_no_match() {
    let out = run_main(
        "boolean present = java.util.stream.LongStream.of(1L, 3L, 5L).filter(n -> n % 2L == 0L).findFirst().isPresent(); System.out.println(present);",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn long_stream_reduce_singleton() {
    let out = run_main(
        "long v = java.util.stream.LongStream.of(99L).reduce((a, b) -> a + b).getAsLong(); System.out.println(v);",
    );
    assert_eq!(out, vec!["99"]);
}

#[test]
fn long_stream_reduce_empty_optional() {
    let out = run_main(
        "boolean present = java.util.stream.LongStream.empty().reduce((a, b) -> a + b).isPresent(); System.out.println(present);",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn long_stream_skip_limit_window() {
    let out = run_main(
        "long s = java.util.stream.LongStream.range(1L, 10L).skip(3L).limit(2L).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn long_stream_generate_constant_sum() {
    let out = run_main(
        "long s = java.util.stream.LongStream.generate(() -> 5L).limit(3L).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn long_stream_range_negative_sum() {
    let out = run_main(
        "long s = java.util.stream.LongStream.rangeClosed(-2L, 2L).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn long_stream_of_max() {
    let out = run_main(
        "long v = java.util.stream.LongStream.of(3L, 9L, 5L).max().getAsLong(); System.out.println(v);",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn long_stream_sequential_sum() {
    let out = run_main(
        "long s = java.util.stream.LongStream.of(1L, 2L, 3L).sequential().sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn long_stream_flat_map_empty() {
    let out = run_main(
        "long c = java.util.stream.LongStream.of(1L, 2L).flatMap(n -> java.util.stream.LongStream.empty()).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn long_stream_to_array_last() {
    let out = run_main(
        "long[] arr = java.util.stream.LongStream.of(4L, 5L, 6L).toArray(); System.out.println(arr[arr.length - 1]);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn long_stream_large_value_preserves() {
    let out = run_main(
        "long v = java.util.stream.LongStream.of(10000000000L).findFirst().getAsLong(); System.out.println(v == 10000000000L);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn long_stream_map_to_int_after_filter() {
    let out = run_main(
        "int s = java.util.stream.LongStream.range(1L, 6L).filter(n -> n > 2L).mapToInt(n -> (int) n).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn long_stream_peek_no_sum_change() {
    let out = run_main(
        "long s = java.util.stream.LongStream.of(2L, 3L).peek(n -> n + 100L).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn long_stream_range_empty_start_equals_end() {
    let out = run_main(
        "long c = java.util.stream.LongStream.range(7L, 7L).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["0"]);
}
