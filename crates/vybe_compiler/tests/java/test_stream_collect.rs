use crate::helpers::run_main;

#[test]
fn collectors_to_list_preserves_stream_element_order() {
    let out = run_main(
        "java.util.List<Integer> out = java.util.Arrays.asList(3, 1, 2).stream().collect(java.util.stream.Collectors.toList()); System.out.println(out.get(0)); System.out.println(out.get(2));",
    );
    assert_eq!(out, vec!["3", "2"]);
}

#[test]
fn collectors_to_list_after_map_collects_transformed_values() {
    let out = run_main(
        "java.util.List<Integer> out = java.util.Arrays.asList(1, 2, 3).stream().map(n -> n + 10).collect(java.util.stream.Collectors.toList()); System.out.println(out.get(0)); System.out.println(out.get(2));",
    );
    assert_eq!(out, vec!["11", "13"]);
}

#[test]
fn collectors_to_list_on_filtered_stream_keeps_matches_only() {
    let out = run_main(
        "java.util.List<Integer> evens = java.util.Arrays.asList(1, 2, 3, 4, 5, 6).stream().filter(n -> n % 2 == 0).collect(java.util.stream.Collectors.toList()); System.out.println(evens.size()); System.out.println(evens.get(1));",
    );
    assert_eq!(out, vec!["3", "6"]);
}

#[test]
fn collectors_to_list_on_empty_stream_yields_empty_list() {
    let out = run_main(
        "java.util.List<Integer> empty = java.util.Arrays.asList(1, 2, 3).stream().filter(n -> n > 10).collect(java.util.stream.Collectors.toList()); System.out.println(empty.size());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn collectors_to_set_removes_duplicate_integers() {
    let out = run_main(
        "java.util.Set<Integer> set = java.util.Arrays.asList(1, 2, 2, 3, 3, 3).stream().collect(java.util.stream.Collectors.toSet()); System.out.println(set.size());",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn collectors_to_set_on_unique_elements_keeps_full_count() {
    let out = run_main(
        "java.util.Set<Integer> set = java.util.Arrays.asList(4, 5, 6).stream().collect(java.util.stream.Collectors.toSet()); System.out.println(set.size());",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn collectors_to_set_after_map_still_deduplicates() {
    let out = run_main(
        "java.util.Set<Integer> set = java.util.Arrays.asList(1, 2, 3, 4).stream().map(n -> n % 2).collect(java.util.stream.Collectors.toSet()); System.out.println(set.size());",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn collectors_to_set_on_strings_collapses_equal_entries() {
    let out = run_main(
        "java.util.Set<String> set = java.util.Arrays.asList(\"a\", \"b\", \"a\").stream().collect(java.util.stream.Collectors.toSet()); System.out.println(set.size());",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn collectors_to_map_builds_key_value_pairs_from_strings() {
    let out = run_main(
        "java.util.Map<String, Integer> map = java.util.Arrays.asList(\"one\", \"two\").stream().collect(java.util.stream.Collectors.toMap(s -> s, String::length)); System.out.println(map.get(\"one\")); System.out.println(map.get(\"two\"));",
    );
    assert_eq!(out, vec!["3", "3"]);
}

#[test]
fn collectors_to_map_uses_identity_for_integer_keys_and_values() {
    let out = run_main(
        "java.util.Map<Integer, Integer> map = java.util.Arrays.asList(1, 2, 3).stream().collect(java.util.stream.Collectors.toMap(n -> n, n -> n)); System.out.println(map.get(2)); System.out.println(map.size());",
    );
    assert_eq!(out, vec!["2", "3"]);
}

#[test]
fn collectors_to_map_after_filter_maps_survivors_only() {
    let out = run_main(
        "java.util.Map<Integer, Integer> map = java.util.Arrays.asList(1, 2, 3, 4).stream().filter(n -> n % 2 == 0).collect(java.util.stream.Collectors.toMap(n -> n, n -> n * 10)); System.out.println(map.get(2)); System.out.println(map.get(4));",
    );
    assert_eq!(out, vec!["20", "40"]);
}

#[test]
fn collectors_to_map_with_value_mapper_squares_keys() {
    let out = run_main(
        "java.util.Map<Integer, Integer> map = java.util.Arrays.asList(2, 3).stream().collect(java.util.stream.Collectors.toMap(n -> n, n -> n * n)); System.out.println(map.get(2)); System.out.println(map.get(3));",
    );
    assert_eq!(out, vec!["4", "9"]);
}

#[test]
fn collectors_joining_concatenates_strings_with_no_delimiter() {
    let out = run_main(
        "String joined = java.util.Arrays.asList(\"a\", \"b\", \"c\").stream().collect(java.util.stream.Collectors.joining()); System.out.println(joined);",
    );
    assert_eq!(out, vec!["abc"]);
}

#[test]
fn collectors_joining_uses_comma_delimiter_between_elements() {
    let out = run_main(
        "String joined = java.util.Arrays.asList(\"x\", \"y\", \"z\").stream().collect(java.util.stream.Collectors.joining(\", \")); System.out.println(joined);",
    );
    assert_eq!(out, vec!["x, y, z"]);
}

#[test]
fn collectors_joining_with_prefix_and_suffix_wraps_result() {
    let out = run_main(
        "String joined = java.util.Arrays.asList(\"1\", \"2\").stream().collect(java.util.stream.Collectors.joining(\",\", \"[\", \"]\")); System.out.println(joined);",
    );
    assert_eq!(out, vec!["[1,2]"]);
}

#[test]
fn collectors_joining_on_mapped_integers_produces_digit_string() {
    let out = run_main(
        "String joined = java.util.Arrays.asList(1, 2, 3).stream().map(n -> \"\" + n).collect(java.util.stream.Collectors.joining(\"\")); System.out.println(joined);",
    );
    assert_eq!(out, vec!["123"]);
}

#[test]
fn collectors_joining_on_empty_stream_yields_empty_string() {
    let out = run_main(
        "String joined = java.util.Arrays.asList(1, 2).stream().filter(n -> n > 5).map(n -> \"\" + n).collect(java.util.stream.Collectors.joining(\",\")); System.out.println(joined.length());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn collectors_counting_returns_element_count_as_long() {
    let out = run_main(
        "Long count = java.util.Arrays.asList(1, 2, 3, 4, 5).stream().collect(java.util.stream.Collectors.counting()); System.out.println(count);",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn collectors_counting_after_filter_counts_matches_only() {
    let out = run_main(
        "Long count = java.util.Arrays.asList(1, 2, 3, 4).stream().filter(n -> n % 2 == 0).collect(java.util.stream.Collectors.counting()); System.out.println(count);",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn collectors_counting_on_empty_stream_is_zero() {
    let out = run_main(
        "Long count = java.util.Arrays.asList(1).stream().filter(n -> n > 9).collect(java.util.stream.Collectors.counting()); System.out.println(count);",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn collectors_counting_after_distinct_counts_unique_items() {
    let out = run_main(
        "Long count = java.util.Arrays.asList(1, 1, 2, 2, 3).stream().distinct().collect(java.util.stream.Collectors.counting()); System.out.println(count);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn collectors_summing_int_adds_all_mapped_values() {
    let out = run_main(
        "Integer sum = java.util.Arrays.asList(1, 2, 3, 4).stream().collect(java.util.stream.Collectors.summingInt(n -> n)); System.out.println(sum);",
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn collectors_summing_int_after_map_doubles_each_contribution() {
    let out = run_main(
        "Integer sum = java.util.Arrays.asList(1, 2, 3).stream().collect(java.util.stream.Collectors.summingInt(n -> n * 2)); System.out.println(sum);",
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn collectors_summing_int_on_filtered_stream_sums_matches() {
    let out = run_main(
        "Integer sum = java.util.Arrays.asList(1, 2, 3, 4, 5).stream().filter(n -> n > 2).collect(java.util.stream.Collectors.summingInt(n -> n)); System.out.println(sum);",
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn collectors_summing_int_on_empty_stream_is_zero() {
    let out = run_main(
        "Integer sum = java.util.Arrays.asList(1, 2).stream().filter(n -> n > 10).collect(java.util.stream.Collectors.summingInt(n -> n)); System.out.println(sum);",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn collectors_averaging_int_computes_mean_of_values() {
    let out = run_main(
        "Double avg = java.util.Arrays.asList(2, 4, 6).stream().collect(java.util.stream.Collectors.averagingInt(n -> n)); System.out.println(avg.intValue());",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn collectors_averaging_int_on_singleton_stream_returns_that_value() {
    let out = run_main(
        "Double avg = java.util.Arrays.asList(7).stream().collect(java.util.stream.Collectors.averagingInt(n -> n)); System.out.println(avg.intValue());",
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn collectors_averaging_int_after_filter_averages_survivors() {
    let out = run_main(
        "Double avg = java.util.Arrays.asList(1, 2, 3, 4).stream().filter(n -> n % 2 == 0).collect(java.util.stream.Collectors.averagingInt(n -> n)); System.out.println(avg.intValue());",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn collectors_grouping_by_length_buckets_strings() {
    let out = run_main(
        "java.util.Map<Integer, java.util.List<String>> groups = java.util.Arrays.asList(\"a\", \"bb\", \"ccc\", \"dd\").stream().collect(java.util.stream.Collectors.groupingBy(s -> s.length())); System.out.println(groups.get(2).size()); System.out.println(groups.get(3).size());",
    );
    assert_eq!(out, vec!["2", "1"]);
}

#[test]
fn collectors_grouping_by_parity_splits_integers_into_two_lists() {
    let out = run_main(
        "java.util.Map<Boolean, java.util.List<Integer>> groups = java.util.Arrays.asList(1, 2, 3, 4).stream().collect(java.util.stream.Collectors.groupingBy(n -> n % 2 == 0)); System.out.println(groups.get(true).size()); System.out.println(groups.get(false).size());",
    );
    assert_eq!(out, vec!["2", "2"]);
}

#[test]
fn collectors_grouping_by_first_character_clusters_words() {
    let out = run_main(
        "java.util.Map<Character, java.util.List<String>> groups = java.util.Arrays.asList(\"ant\", \"ape\", \"bat\").stream().collect(java.util.stream.Collectors.groupingBy(s -> s.charAt(0))); System.out.println(groups.get('a').size()); System.out.println(groups.get('b').size());",
    );
    assert_eq!(out, vec!["2", "1"]);
}

#[test]
fn collectors_grouping_by_with_counting_downstream_counts_per_bucket() {
    let out = run_main(
        "java.util.Map<Integer, Long> counts = java.util.Arrays.asList(1, 2, 3, 4, 5, 6).stream().collect(java.util.stream.Collectors.groupingBy(n -> n % 3, java.util.stream.Collectors.counting())); System.out.println(counts.get(0)); System.out.println(counts.get(1));",
    );
    assert_eq!(out, vec!["2", "2"]);
}

#[test]
fn collectors_partitioning_by_splits_even_and_odd_integers() {
    let out = run_main(
        "java.util.Map<Boolean, java.util.List<Integer>> parts = java.util.Arrays.asList(1, 2, 3, 4, 5).stream().collect(java.util.stream.Collectors.partitioningBy(n -> n % 2 == 0)); System.out.println(parts.get(true).size()); System.out.println(parts.get(false).size());",
    );
    assert_eq!(out, vec!["2", "3"]);
}

#[test]
fn collectors_partitioning_by_on_strings_uses_length_threshold() {
    let out = run_main(
        "java.util.Map<Boolean, java.util.List<String>> parts = java.util.Arrays.asList(\"a\", \"bb\", \"ccc\").stream().collect(java.util.stream.Collectors.partitioningBy(s -> s.length() > 1)); System.out.println(parts.get(true).size()); System.out.println(parts.get(false).size());",
    );
    assert_eq!(out, vec!["2", "1"]);
}

#[test]
fn collectors_partitioning_by_with_counting_downstream() {
    let out = run_main(
        "java.util.Map<Boolean, Long> counts = java.util.Arrays.asList(1, 2, 3, 4).stream().collect(java.util.stream.Collectors.partitioningBy(n -> n > 2, java.util.stream.Collectors.counting())); System.out.println(counts.get(true)); System.out.println(counts.get(false));",
    );
    assert_eq!(out, vec!["2", "2"]);
}

#[test]
fn collectors_partitioning_by_always_creates_two_buckets() {
    let out = run_main(
        "java.util.Map<Boolean, java.util.List<Integer>> parts = java.util.Arrays.asList(2, 4, 6).stream().collect(java.util.stream.Collectors.partitioningBy(n -> n % 2 == 0)); System.out.println(parts.get(true).size()); System.out.println(parts.get(false).size());",
    );
    assert_eq!(out, vec!["3", "0"]);
}

#[test]
fn collectors_mapping_transforms_elements_before_list_downstream() {
    let out = run_main(
        "java.util.List<Integer> lengths = java.util.Arrays.asList(\"a\", \"bb\", \"ccc\").stream().collect(java.util.stream.Collectors.mapping(String::length, java.util.stream.Collectors.toList())); System.out.println(lengths.get(0)); System.out.println(lengths.get(2));",
    );
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn collectors_mapping_inside_grouping_by_stores_doubled_values() {
    let out = run_main(
        "java.util.Map<Boolean, java.util.List<Integer>> groups = java.util.Arrays.asList(1, 2, 3, 4).stream().collect(java.util.stream.Collectors.groupingBy(n -> n % 2 == 0, java.util.stream.Collectors.mapping(n -> n * 10, java.util.stream.Collectors.toList()))); System.out.println(groups.get(false).get(0)); System.out.println(groups.get(true).get(1));",
    );
    assert_eq!(out, vec!["10", "40"]);
}

#[test]
fn collectors_mapping_to_set_downstream_deduplicates_mapped_results() {
    let out = run_main(
        "java.util.Set<Integer> mods = java.util.Arrays.asList(1, 2, 3, 4, 5).stream().collect(java.util.stream.Collectors.mapping(n -> n % 3, java.util.stream.Collectors.toSet())); System.out.println(mods.size());",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn collectors_filtering_applies_predicate_before_list_collection() {
    let out = run_main(
        "java.util.List<Integer> out = java.util.Arrays.asList(1, 2, 3, 4, 5).stream().collect(java.util.stream.Collectors.filtering(n -> n > 2, java.util.stream.Collectors.toList())); System.out.println(out.size()); System.out.println(out.get(0));",
    );
    assert_eq!(out, vec!["3", "3"]);
}

#[test]
fn collectors_filtering_with_counting_counts_only_matches() {
    let out = run_main(
        "Long count = java.util.Arrays.asList(1, 2, 3, 4, 5, 6).stream().collect(java.util.stream.Collectors.filtering(n -> n % 2 == 0, java.util.stream.Collectors.counting())); System.out.println(count);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn collectors_filtering_inside_grouping_by_prunes_each_bucket() {
    let out = run_main(
        "java.util.Map<Boolean, java.util.List<Integer>> groups = java.util.Arrays.asList(1, 2, 3, 4, 5, 6).stream().collect(java.util.stream.Collectors.groupingBy(n -> n % 2 == 0, java.util.stream.Collectors.filtering(n -> n > 2, java.util.stream.Collectors.toList()))); System.out.println(groups.get(true).size()); System.out.println(groups.get(false).size());",
    );
    assert_eq!(out, vec!["2", "1"]);
}

#[test]
fn collectors_to_list_and_counting_agree_on_same_stream_size() {
    let out = run_main(
        "java.util.List<Integer> list = java.util.Arrays.asList(9, 8, 7).stream().collect(java.util.stream.Collectors.toList()); Long count = java.util.Arrays.asList(9, 8, 7).stream().collect(java.util.stream.Collectors.counting()); System.out.println(list.size()); System.out.println(count);",
    );
    assert_eq!(out, vec!["3", "3"]);
}

#[test]
fn collectors_joining_after_sorted_stream_orders_output() {
    let out = run_main(
        "String joined = java.util.Arrays.asList(\"c\", \"a\", \"b\").stream().sorted().collect(java.util.stream.Collectors.joining(\"\")); System.out.println(joined);",
    );
    assert_eq!(out, vec!["abc"]);
}

#[test]
fn collectors_summing_int_matches_reduce_sum_on_same_data() {
    let out = run_main(
        "Integer collected = java.util.Arrays.asList(1, 2, 3).stream().collect(java.util.stream.Collectors.summingInt(n -> n)); int reduced = java.util.Arrays.asList(1, 2, 3).stream().reduce(0, (a, b) -> a + b); System.out.println(collected); System.out.println(reduced);",
    );
    assert_eq!(out, vec!["6", "6"]);
}
