use crate::helpers::run_main;

#[test]
fn arraylist_spliterator_estimate_size_matches_list() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); list.add(3); System.out.println(list.spliterator().estimateSize());"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn arraylist_spliterator_try_advance_reads_first() {
    let out = run_main(
        r#"java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add("first"); java.util.Spliterator<String> sp = list.spliterator(); final String[] box = new String[1]; sp.tryAdvance(x -> box[0] = x); System.out.println(box[0]);"#,
    );
    assert_eq!(out, vec!["first"]);
}

#[test]
fn arraylist_spliterator_try_advance_false_when_empty() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); System.out.println(list.spliterator().tryAdvance(x -> {}));"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn arraylist_spliterator_for_each_remaining_visits_all() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); list.add(3); final int[] sum = new int[1]; list.spliterator().forEachRemaining(x -> sum[0] += x); System.out.println(sum[0]);"#,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn arraylist_spliterator_characteristics_includes_ordered() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); int ch = list.spliterator().characteristics(); System.out.println((ch & java.util.Spliterator.ORDERED) != 0);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn arraylist_spliterator_characteristics_includes_sized() {
    let out = run_main(
        r#"java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add("a"); int ch = list.spliterator().characteristics(); System.out.println((ch & java.util.Spliterator.SIZED) != 0);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn arraylist_spliterator_characteristics_includes_subsized() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); int ch = list.spliterator().characteristics(); System.out.println((ch & java.util.Spliterator.SUBSIZED) != 0);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn arraylist_spliterator_try_split_produces_sub_spliterator() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); list.add(3); list.add(4); java.util.Spliterator<Integer> sp = list.spliterator(); java.util.Spliterator<Integer> sub = sp.trySplit(); System.out.println(sub != null);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn arraylist_spliterator_after_split_both_have_elements() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); for (int i = 1; i <= 4; i++) list.add(i); java.util.Spliterator<Integer> sp = list.spliterator(); java.util.Spliterator<Integer> sub = sp.trySplit(); final int[] c1 = new int[1]; final int[] c2 = new int[1]; if (sub != null) sub.forEachRemaining(x -> c1[0]++); sp.forEachRemaining(x -> c2[0]++); System.out.println(c1[0] + c2[0]);"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn arraylist_spliterator_get_comparator_null_for_natural_order() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); System.out.println(list.spliterator().getComparator() == null);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn arraylist_spliterator_on_empty_list_size_zero() {
    let out = run_main(
        r#"java.util.ArrayList<String> list = new java.util.ArrayList<String>(); System.out.println(list.spliterator().estimateSize());"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn arraylist_spliterator_single_element_try_split_null() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(99); System.out.println(list.spliterator().trySplit() == null);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn arraylist_spliterator_try_advance_twice_reads_both() {
    let out = run_main(
        r#"java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add("a"); list.add("b"); java.util.Spliterator<String> sp = list.spliterator(); final String[] first = new String[1]; final String[] second = new String[1]; sp.tryAdvance(x -> first[0] = x); sp.tryAdvance(x -> second[0] = x); System.out.println(first[0] + second[0]);"#,
    );
    assert_eq!(out, vec!["ab"]);
}

#[test]
fn arraylist_spliterator_for_each_remaining_on_empty_noop() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); final int[] n = new int[1]; list.spliterator().forEachRemaining(x -> n[0]++); System.out.println(n[0]);"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn arraylist_spliterator_characteristics_nonnull_for_nonempty() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); System.out.println(list.spliterator().characteristics() > 0);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn arraylist_spliterator_ordered_traversal_preserves_insertion() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(10); list.add(20); final int[] first = new int[1]; list.spliterator().tryAdvance(x -> first[0] = x); System.out.println(first[0]);"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn arraylist_spliterator_after_partial_advance_remaining_count() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); list.add(3); java.util.Spliterator<Integer> sp = list.spliterator(); sp.tryAdvance(x -> {}); final int[] n = new int[1]; sp.forEachRemaining(x -> n[0]++); System.out.println(n[0]);"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn arraylist_spliterator_on_list_of_strings_concat() {
    let out = run_main(
        r#"java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add("x"); list.add("y"); final StringBuilder sb = new StringBuilder(); list.spliterator().forEachRemaining(sb::append); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["xy"]);
}

#[test]
fn arraylist_spliterator_split_then_advance_on_both_halves() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); java.util.Spliterator<Integer> sp = list.spliterator(); java.util.Spliterator<Integer> sub = sp.trySplit(); final int[] total = new int[1]; if (sub != null) sub.forEachRemaining(x -> total[0] += x); sp.forEachRemaining(x -> total[0] += x); System.out.println(total[0]);"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn arraylist_spliterator_has_immutable_characteristic_false() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); int ch = list.spliterator().characteristics(); System.out.println((ch & java.util.Spliterator.IMMUTABLE) != 0);"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn arraylist_spliterator_has_concurrent_characteristic_false() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); int ch = list.spliterator().characteristics(); System.out.println((ch & java.util.Spliterator.CONCURRENT) != 0);"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn arraylist_spliterator_has_sorted_characteristic_false() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(3); list.add(1); int ch = list.spliterator().characteristics(); System.out.println((ch & java.util.Spliterator.SORTED) != 0);"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn arraylist_spliterator_has_distinct_characteristic_false() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(1); int ch = list.spliterator().characteristics(); System.out.println((ch & java.util.Spliterator.DISTINCT) != 0);"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn arraylist_spliterator_try_advance_returns_false_at_end() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); java.util.Spliterator<Integer> sp = list.spliterator(); sp.tryAdvance(x -> {}); System.out.println(sp.tryAdvance(x -> {}));"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn arraylist_spliterator_large_list_estimate_size() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); for (int i = 0; i < 100; i++) list.add(i); System.out.println(list.spliterator().estimateSize());"#,
    );
    assert_eq!(out, vec!["100"]);
}

#[test]
fn arraylist_spliterator_for_each_remaining_with_null_element() {
    let out = run_main(
        r#"java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(null); list.add("x"); final int[] nulls = new int[1]; list.spliterator().forEachRemaining(x -> { if (x == null) nulls[0]++; }); System.out.println(nulls[0]);"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn arraylist_spliterator_multiple_splits_cover_all_elements() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); for (int i = 0; i < 8; i++) list.add(i); java.util.Spliterator<Integer> sp = list.spliterator(); final int[] sum = new int[1]; while (true) { java.util.Spliterator<Integer> sub = sp.trySplit(); if (sub == null) { sp.forEachRemaining(x -> sum[0] += x); break; } sub.forEachRemaining(x -> sum[0] += x); } System.out.println(sum[0]);"#,
    );
    assert_eq!(out, vec!["28"]);
}

#[test]
fn arraylist_spliterator_stream_from_spliterator_count() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); list.add(3); long count = java.util.stream.StreamSupport.stream(list.spliterator(), false).count(); System.out.println(count);"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn arraylist_spliterator_stream_parallel_flag_false() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); java.util.stream.Stream<Integer> s = java.util.stream.StreamSupport.stream(list.spliterator(), false); System.out.println(s.isParallel());"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn arraylist_spliterator_stream_parallel_flag_true() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); java.util.stream.Stream<Integer> s = java.util.stream.StreamSupport.stream(list.spliterator(), true); System.out.println(s.isParallel());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn arraylist_spliterator_reduce_sum_via_stream() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(4); list.add(5); int sum = java.util.stream.StreamSupport.stream(list.spliterator(), false).reduce(0, Integer::sum); System.out.println(sum);"#,
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn arraylist_spliterator_filter_via_stream() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); list.add(3); long count = java.util.stream.StreamSupport.stream(list.spliterator(), false).filter(x -> x > 1).count(); System.out.println(count);"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn arraylist_spliterator_map_to_string_via_stream() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(7); String s = java.util.stream.StreamSupport.stream(list.spliterator(), false).map(String::valueOf).findFirst().get(); System.out.println(s);"#,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn arraylist_spliterator_collect_to_list_via_stream() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); java.util.List<Integer> copy = java.util.stream.StreamSupport.stream(list.spliterator(), false).collect(java.util.stream.Collectors.toList()); System.out.println(copy.size());"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn arraylist_spliterator_nonnull_characteristic_true() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); int ch = list.spliterator().characteristics(); System.out.println((ch & java.util.Spliterator.NONNULL) != 0);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn arraylist_spliterator_after_clear_estimate_zero() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.clear(); System.out.println(list.spliterator().estimateSize());"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn arraylist_spliterator_try_split_on_two_element_list() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); java.util.Spliterator<Integer> sub = list.spliterator().trySplit(); System.out.println(sub != null);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn arraylist_spliterator_for_each_remaining_stops_after_all_consumed() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); java.util.Spliterator<Integer> sp = list.spliterator(); sp.forEachRemaining(x -> {}); System.out.println(sp.tryAdvance(x -> {}));"#,
    );
    assert_eq!(out, vec!["false"]);
}
