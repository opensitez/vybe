use crate::helpers::run_main;

#[test]
fn double_stream_of_three_sums() {
    let out = run_main(
        "double s = java.util.stream.DoubleStream.of(1.0, 2.0, 3.0).sum(); System.out.println((int) s);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn double_stream_of_single_count() {
    let out =
        run_main("long c = java.util.stream.DoubleStream.of(42.5).count(); System.out.println(c);");
    assert_eq!(out, vec!["1"]);
}

#[test]
fn double_stream_empty_count() {
    let out =
        run_main("long c = java.util.stream.DoubleStream.empty().count(); System.out.println(c);");
    assert_eq!(out, vec!["0"]);
}

#[test]
fn double_stream_iterate_halves() {
    let out = run_main(
        "double last = java.util.stream.DoubleStream.iterate(8.0, n -> n / 2.0).limit(3).reduce((a, b) -> b).getAsDouble(); System.out.println((int) last);",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn double_stream_iterate_with_predicate() {
    let out = run_main(
        "long c = java.util.stream.DoubleStream.iterate(1.0, n -> n < 10.0, n -> n + 3.0).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn double_stream_generate_increments() {
    let out = run_main(
        "double[] box = {0.0}; double s = java.util.stream.DoubleStream.generate(() -> ++box[0]).limit(3).sum(); System.out.println((int) s);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn double_stream_filter_above_two() {
    let out = run_main(
        "double s = java.util.stream.DoubleStream.of(1.0, 2.5, 3.0, 4.0).filter(n -> n > 2.0).sum(); System.out.println((int) s);",
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn double_stream_filter_none_match() {
    let out = run_main(
        "long c = java.util.stream.DoubleStream.of(1.0, 2.0).filter(n -> n > 10.0).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn double_stream_map_doubles() {
    let out = run_main(
        "double s = java.util.stream.DoubleStream.of(1.5, 2.0).map(n -> n * 2.0).sum(); System.out.println((int) s);",
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn double_stream_map_to_int_truncates() {
    let out = run_main(
        "int s = java.util.stream.DoubleStream.of(2.9, 3.1).mapToInt(n -> (int) n).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn double_stream_map_to_long_rounds() {
    let out = run_main(
        "long s = java.util.stream.DoubleStream.of(2.5, 3.5).mapToLong(n -> (long) n).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn double_stream_map_to_obj_count() {
    let out = run_main(
        "long c = java.util.stream.DoubleStream.of(1.0, 2.0).mapToObj(n -> \"D\" + n).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn double_stream_flat_map_duplicates() {
    let out = run_main(
        "double s = java.util.stream.DoubleStream.of(1.0, 2.0).flatMap(n -> java.util.stream.DoubleStream.of(n, n)).sum(); System.out.println((int) s);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn double_stream_distinct_on_duplicates() {
    let out = run_main(
        "long c = java.util.stream.DoubleStream.of(1.0, 1.0, 2.0, 2.0).distinct().count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn double_stream_sorted_desc_via_boxed() {
    let out = run_main(
        "double first = java.util.stream.DoubleStream.of(3.0, 1.0, 2.0).boxed().sorted((a, b) -> Double.compare(b, a)).findFirst().get(); System.out.println((int) first);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn double_stream_peek_counts() {
    let out = run_main(
        "double[] seen = {0.0}; double s = java.util.stream.DoubleStream.of(1.0, 2.0, 3.0).peek(n -> seen[0]++).sum(); System.out.println((int) seen[0]); System.out.println((int) s);",
    );
    assert_eq!(out, vec!["3", "6"]);
}

#[test]
fn double_stream_limit_iterate() {
    let out = run_main(
        "long c = java.util.stream.DoubleStream.iterate(0.0, n -> n + 1.0).limit(5).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn double_stream_skip_prefix() {
    let out = run_main(
        "double s = java.util.stream.DoubleStream.of(1.0, 2.0, 3.0, 4.0, 5.0).skip(2).sum(); System.out.println((int) s);",
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn double_stream_for_each_prints() {
    let out = run_main(
        "java.util.stream.DoubleStream.of(1.5, 2.5).forEach(n -> System.out.println((int) n));",
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn double_stream_to_array_length() {
    let out = run_main(
        "int len = java.util.stream.DoubleStream.of(1.0, 2.0, 3.0).toArray().length; System.out.println(len);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn double_stream_reduce_with_identity() {
    let out = run_main(
        "double s = java.util.stream.DoubleStream.of(1.0, 2.0, 3.0).reduce(0.0, (a, b) -> a + b); System.out.println((int) s);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn double_stream_reduce_without_identity() {
    let out = run_main(
        "double p = java.util.stream.DoubleStream.of(2.0, 3.0, 4.0).reduce((a, b) -> a * b).getAsDouble(); System.out.println((int) p);",
    );
    assert_eq!(out, vec!["24"]);
}

#[test]
fn double_stream_min_value() {
    let out = run_main(
        "double v = java.util.stream.DoubleStream.of(7.0, 2.0, 9.0).min().getAsDouble(); System.out.println((int) v);",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn double_stream_max_value() {
    let out = run_main(
        "double v = java.util.stream.DoubleStream.of(7.0, 2.0, 9.0).max().getAsDouble(); System.out.println((int) v);",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn double_stream_average_triplet() {
    let out = run_main(
        "double a = java.util.stream.DoubleStream.of(2.0, 5.0, 8.0).average().getAsDouble(); System.out.println((int) a);",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn double_stream_sum_of_squares() {
    let out = run_main(
        "double s = java.util.stream.DoubleStream.of(1.0, 2.0, 3.0).map(n -> n * n).sum(); System.out.println((int) s);",
    );
    assert_eq!(out, vec!["14"]);
}

#[test]
fn double_stream_count_above_one() {
    let out = run_main(
        "long c = java.util.stream.DoubleStream.of(0.5, 1.5, 2.5, 3.5).filter(n -> n > 1.0).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn double_stream_any_match_true() {
    let out = run_main(
        "boolean b = java.util.stream.DoubleStream.of(1.0, 2.0, 6.0).anyMatch(n -> n > 5.0); System.out.println(b);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn double_stream_any_match_false_empty() {
    let out = run_main(
        "boolean b = java.util.stream.DoubleStream.empty().anyMatch(n -> n > 0.0); System.out.println(b);",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn double_stream_all_match_positive() {
    let out = run_main(
        "boolean b = java.util.stream.DoubleStream.of(1.0, 2.0, 3.0).allMatch(n -> n > 0.0); System.out.println(b);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn double_stream_all_match_false() {
    let out = run_main(
        "boolean b = java.util.stream.DoubleStream.of(1.0, -1.0, 3.0).allMatch(n -> n > 0.0); System.out.println(b);",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn double_stream_none_match_negative() {
    let out = run_main(
        "boolean b = java.util.stream.DoubleStream.of(-1.0, -2.0).noneMatch(n -> n > 0.0); System.out.println(b);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn double_stream_none_match_false() {
    let out = run_main(
        "boolean b = java.util.stream.DoubleStream.of(-1.0, 2.0).noneMatch(n -> n > 0.0); System.out.println(b);",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn double_stream_find_first_above_four() {
    let out = run_main(
        "double v = java.util.stream.DoubleStream.of(1.0, 3.0, 5.0, 6.0).filter(n -> n > 4.0).findFirst().getAsDouble(); System.out.println((int) v);",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn double_stream_find_any_less_than_nine() {
    let out = run_main(
        "double v = java.util.stream.DoubleStream.of(9.0, 8.0, 7.0).filter(n -> n < 9.0).findAny().getAsDouble(); System.out.println((int) v);",
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn double_stream_boxed_list_size() {
    let out = run_main(
        "int sz = java.util.stream.DoubleStream.of(1.0, 2.0, 3.0).boxed().toList().size(); System.out.println(sz);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn double_stream_concat_two() {
    let out = run_main(
        "long c = java.util.stream.DoubleStream.concat(java.util.stream.DoubleStream.of(1.0, 2.0), java.util.stream.DoubleStream.of(3.0)).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn double_stream_builder_sum() {
    let out = run_main(
        "double s = java.util.stream.DoubleStream.builder().add(2.0).add(3.0).build().sum(); System.out.println((int) s);",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn double_stream_summary_statistics_sum() {
    let out = run_main(
        "double s = java.util.stream.DoubleStream.of(1.0, 2.0, 3.0).summaryStatistics().getSum(); System.out.println((int) s);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn double_stream_summary_statistics_max() {
    let out = run_main(
        "double v = java.util.stream.DoubleStream.of(1.0, 5.0, 3.0).summaryStatistics().getMax(); System.out.println((int) v);",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn double_stream_summary_statistics_avg() {
    let out = run_main(
        "double a = java.util.stream.DoubleStream.of(2.0, 4.0, 6.0).summaryStatistics().getAverage(); System.out.println((int) a);",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn double_stream_summary_statistics_count() {
    let out = run_main(
        "long c = java.util.stream.DoubleStream.of(1.0, 2.0, 3.0, 4.0).summaryStatistics().getCount(); System.out.println(c);",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn double_stream_collect_joining() {
    let out = run_main(
        "String j = java.util.stream.DoubleStream.of(1.0, 2.0).mapToObj(n -> \"\" + (int) n).collect(java.util.stream.Collectors.joining(\"\")); System.out.println(j);",
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn double_stream_drop_while_below_three() {
    let out = run_main(
        "double s = java.util.stream.DoubleStream.of(1.0, 2.0, 3.0, 1.0).dropWhile(n -> n < 3.0).sum(); System.out.println((int) s);",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn double_stream_take_while_below_three() {
    let out = run_main(
        "double s = java.util.stream.DoubleStream.of(1.0, 2.0, 3.0, 4.0).takeWhile(n -> n < 3.0).sum(); System.out.println((int) s);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn double_stream_reduce_multiplies_chain() {
    let out = run_main(
        "double p = java.util.stream.DoubleStream.of(2.0, 3.0, 0.5).reduce(1.0, (a, b) -> a * b); System.out.println((int) p);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn double_stream_map_abs_negative() {
    let out = run_main(
        "double s = java.util.stream.DoubleStream.of(-2.0, -1.0, 3.0).map(Math::abs).sum(); System.out.println((int) s);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn double_stream_filter_map_chain() {
    let out = run_main(
        "double s = java.util.stream.DoubleStream.of(1.0, 2.0, 3.0, 4.0, 5.0).filter(n -> n <= 4.0).map(n -> n + 1.0).sum(); System.out.println((int) s);",
    );
    assert_eq!(out, vec!["14"]);
}

#[test]
fn double_stream_distinct_constant() {
    let out = run_main(
        "long c = java.util.stream.DoubleStream.generate(() -> 7.0).limit(4).distinct().count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn double_stream_find_first_no_match() {
    let out = run_main(
        "boolean present = java.util.stream.DoubleStream.of(1.0, 3.0).filter(n -> n > 5.0).findFirst().isPresent(); System.out.println(present);",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn double_stream_reduce_singleton() {
    let out = run_main(
        "double v = java.util.stream.DoubleStream.of(99.0).reduce((a, b) -> a + b).getAsDouble(); System.out.println((int) v);",
    );
    assert_eq!(out, vec!["99"]);
}

#[test]
fn double_stream_reduce_empty() {
    let out = run_main(
        "boolean present = java.util.stream.DoubleStream.empty().reduce((a, b) -> a + b).isPresent(); System.out.println(present);",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn double_stream_skip_limit_window() {
    let out = run_main(
        "double s = java.util.stream.DoubleStream.iterate(1.0, n -> n + 1.0).skip(3).limit(2).sum(); System.out.println((int) s);",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn double_stream_generate_constant() {
    let out = run_main(
        "double s = java.util.stream.DoubleStream.generate(() -> 5.0).limit(3).sum(); System.out.println((int) s);",
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn double_stream_of_max() {
    let out = run_main(
        "double v = java.util.stream.DoubleStream.of(3.0, 9.0, 5.0).max().getAsDouble(); System.out.println((int) v);",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn double_stream_sequential_sum() {
    let out = run_main(
        "double s = java.util.stream.DoubleStream.of(1.0, 2.0, 3.0).sequential().sum(); System.out.println((int) s);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn double_stream_flat_map_empty() {
    let out = run_main(
        "long c = java.util.stream.DoubleStream.of(1.0, 2.0).flatMap(n -> java.util.stream.DoubleStream.empty()).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn double_stream_to_array_last() {
    let out = run_main(
        "double[] arr = java.util.stream.DoubleStream.of(4.0, 5.0, 6.0).toArray(); System.out.println((int) arr[arr.length - 1]);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn double_stream_fractional_average() {
    let out = run_main(
        "double a = java.util.stream.DoubleStream.of(1.0, 2.0).average().getAsDouble(); System.out.println(a == 1.5);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn double_stream_map_to_int_floor() {
    let out = run_main(
        "int s = java.util.stream.DoubleStream.of(2.9, 3.9).mapToInt(n -> (int) Math.floor(n)).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn double_stream_peek_no_sum_change() {
    let out = run_main(
        "double s = java.util.stream.DoubleStream.of(2.0, 3.0).peek(n -> n + 100.0).sum(); System.out.println((int) s);",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn double_stream_range_not_available_use_iterate() {
    let out = run_main(
        "double s = java.util.stream.DoubleStream.iterate(1.0, n -> n + 1.0).limit(5).sum(); System.out.println((int) s);",
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn double_stream_of_spread_min() {
    let out = run_main(
        "double v = java.util.stream.DoubleStream.of(3.0, 9.0, 5.0).min().getAsDouble(); System.out.println((int) v);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn double_stream_match_short_circuit() {
    let out = run_main(
        "boolean b = java.util.stream.DoubleStream.of(1.0, 2.0, 3.0).anyMatch(n -> n == 2.0); System.out.println(b);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn double_stream_map_to_long_after_filter() {
    let out = run_main(
        "long s = java.util.stream.DoubleStream.of(1.0, 2.5, 3.5, 4.5).filter(n -> n > 2.0).mapToLong(n -> (long) n).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn double_stream_zero_values_sum() {
    let out = run_main(
        "double s = java.util.stream.DoubleStream.of(0.0, 0.0, 0.0).sum(); System.out.println((int) s);",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn double_stream_negative_values_sum() {
    let out = run_main(
        "double s = java.util.stream.DoubleStream.of(-1.0, -2.0, 3.0).sum(); System.out.println((int) s);",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn double_stream_collect_to_list_via_boxed() {
    let out = run_main(
        "int sz = java.util.stream.DoubleStream.of(1.1, 2.2).boxed().collect(java.util.stream.Collectors.toList()).size(); System.out.println(sz);",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn double_stream_builder_add_multiple() {
    let out = run_main(
        "double s = java.util.stream.DoubleStream.builder().add(1.0).add(2.0).add(3.0).build().sum(); System.out.println((int) s);",
    );
    assert_eq!(out, vec!["6"]);
}
