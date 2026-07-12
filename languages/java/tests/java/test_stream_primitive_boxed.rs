use crate::helpers::run_main;

#[test]
fn boxed_int_stream_from_list_map_to_int() {
    let out = run_main(
        "int s = java.util.Arrays.asList(1, 2, 3).stream().mapToInt(Integer::intValue).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn boxed_long_stream_from_list_map_to_long() {
    let out = run_main(
        "long s = java.util.Arrays.asList(1L, 2L).stream().mapToLong(Long::longValue).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn boxed_double_stream_from_list_map_to_double() {
    let out = run_main(
        "double s = java.util.Arrays.asList(1.5, 2.5).stream().mapToDouble(Double::doubleValue).sum(); System.out.println((int) s);",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn int_stream_boxed_back_to_stream_count() {
    let out = run_main(
        "long c = java.util.stream.IntStream.of(1, 2, 3).boxed().stream().count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn long_stream_boxed_collects_strings() {
    let out = run_main(
        "String j = java.util.stream.LongStream.of(1L, 2L).boxed().map(l -> \"L\" + l).collect(java.util.stream.Collectors.joining()); System.out.println(j);",
    );
    assert_eq!(out, vec!["L1L2"]);
}

#[test]
fn double_stream_boxed_sorted_desc() {
    let out = run_main(
        "double first = java.util.stream.DoubleStream.of(3.0, 1.0, 2.0).boxed().sorted((a, b) -> Double.compare(b, a)).findFirst().get(); System.out.println((int) first);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn list_stream_map_to_int_unboxes_null_safe() {
    let out = run_main(
        "int s = java.util.Arrays.asList(Integer.valueOf(5), Integer.valueOf(6)).stream().mapToInt(n -> n).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["11"]);
}

#[test]
fn int_stream_boxed_to_list_get() {
    let out = run_main(
        "int v = java.util.stream.IntStream.range(1, 4).boxed().toList().get(1); System.out.println(v);",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn arrays_stream_ints_map_to_obj() {
    let out = run_main(
        "long c = java.util.Arrays.stream(new int[]{1, 2}).mapToObj(n -> n).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn arrays_stream_longs_map_to_double() {
    let out = run_main(
        "double s = java.util.Arrays.stream(new long[]{2L, 4L}).mapToDouble(n -> (double) n).sum(); System.out.println((int) s);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn arrays_stream_doubles_map_to_int() {
    let out = run_main(
        "int s = java.util.Arrays.stream(new double[]{2.9, 3.1}).mapToInt(n -> (int) n).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn stream_of_integers_map_to_int_sum() {
    let out = run_main(
        "int s = java.util.stream.Stream.of(1, 2, 3).mapToInt(Integer::intValue).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn int_stream_as_long_then_boxed_count() {
    let out = run_main(
        "long c = java.util.stream.IntStream.of(1, 2).asLongStream().boxed().count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn int_stream_as_double_then_sum() {
    let out = run_main(
        "double s = java.util.stream.IntStream.of(1, 2, 3).asDoubleStream().sum(); System.out.println((int) s);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn long_stream_map_to_int_sum() {
    let out = run_main(
        "int s = java.util.stream.LongStream.of(10L, 20L).mapToInt(n -> (int) n).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["30"]);
}

#[test]
fn double_stream_map_to_long_sum() {
    let out = run_main(
        "long s = java.util.stream.DoubleStream.of(2.0, 3.0).mapToLong(n -> (long) n).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn boxed_stream_collect_to_set() {
    let out = run_main(
        "int sz = java.util.stream.IntStream.of(1, 2, 2).boxed().collect(java.util.stream.Collectors.toSet()).size(); System.out.println(sz);",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn primitive_concat_boxed_count() {
    let out = run_main(
        "long c = java.util.stream.Stream.concat(java.util.stream.IntStream.of(1).boxed(), java.util.stream.IntStream.of(2).boxed()).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn list_parallel_stream_map_to_int() {
    let out = run_main(
        "int s = java.util.Arrays.asList(1, 2, 3).parallelStream().mapToInt(n -> n).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn optional_int_stream_from_boxed() {
    let out = run_main(
        "int s = java.util.Optional.of(4).stream().mapToInt(n -> n).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn int_stream_map_to_obj_then_stream_count() {
    let out = run_main(
        "long c = java.util.stream.IntStream.of(1, 2).mapToObj(n -> n).flatMap(n -> java.util.stream.Stream.of(n, n)).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn double_list_stream_boxed_average() {
    let out = run_main(
        "double a = java.util.Arrays.asList(2.0, 4.0).stream().mapToDouble(d -> d).average().getAsDouble(); System.out.println((int) a);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn int_stream_builder_boxed_join() {
    let out = run_main(
        "String j = java.util.stream.IntStream.builder().add(1).add(2).build().boxed().map(n -> \"\" + n).collect(java.util.stream.Collectors.joining(\"-\")); System.out.println(j);",
    );
    assert_eq!(out, vec!["1-2"]);
}

#[test]
fn long_stream_boxed_filter_count() {
    let out = run_main(
        "long c = java.util.stream.LongStream.of(1L, 2L, 3L).boxed().filter(n -> n > 1L).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn double_stream_boxed_map_to_int_sum() {
    let out = run_main(
        "int s = java.util.stream.DoubleStream.of(1.9, 2.1).boxed().mapToInt(n -> n.intValue()).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["3"]);
}
