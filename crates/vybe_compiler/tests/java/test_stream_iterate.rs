use crate::helpers::run_main;

#[test]
fn stream_iterate_seed_only_limit() {
    let out = run_main(
        "java.util.List<Integer> list = java.util.stream.Stream.iterate(1, n -> n + 1).limit(3).toList(); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn stream_iterate_with_predicate_stops() {
    let out = run_main(
        "long c = java.util.stream.Stream.iterate(1, n -> n < 10, n -> n + 3).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn stream_iterate_doubles_seed() {
    let out = run_main(
        "int s = java.util.stream.Stream.iterate(2, n -> n * 2).limit(4).mapToInt(n -> n).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["30"]);
}

#[test]
fn stream_iterate_fibonacci_style() {
    let out = run_main(
        "int[] p = {0, 1}; int v = java.util.stream.Stream.iterate(0, n -> { int nx = p[0] + p[1]; p[0] = p[1]; p[1] = nx; return nx; }).limit(5).reduce((a, b) -> b).get(); System.out.println(v);",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn stream_iterate_limit_one() {
    let out = run_main(
        "int v = java.util.stream.Stream.iterate(9, n -> n + 1).limit(1).findFirst().get(); System.out.println(v);",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn stream_iterate_filter_evens() {
    let out = run_main(
        "long c = java.util.stream.Stream.iterate(1, n -> n + 1).filter(n -> n % 2 == 0).limit(2).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn stream_iterate_map_to_string() {
    let out = run_main(
        "String j = java.util.stream.Stream.iterate(1, n -> n + 1).limit(3).map(n -> \"n\" + n).collect(java.util.stream.Collectors.joining()); System.out.println(j);",
    );
    assert_eq!(out, vec!["n1n2n3"]);
}

#[test]
fn stream_iterate_skip_then_limit() {
    let out = run_main(
        "int s = java.util.stream.Stream.iterate(1, n -> n + 1).skip(2).limit(2).mapToInt(n -> n).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn stream_iterate_take_while_below_five() {
    let out = run_main(
        "int s = java.util.stream.Stream.iterate(1, n -> n + 1).takeWhile(n -> n < 5).mapToInt(n -> n).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn stream_iterate_drop_while_below_three() {
    let out = run_main(
        "int s = java.util.stream.Stream.iterate(1, n -> n + 1).dropWhile(n -> n < 3).limit(2).mapToInt(n -> n).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn stream_iterate_peek_side_effect() {
    let out = run_main(
        "int[] c = {0}; java.util.stream.Stream.iterate(1, n -> n + 1).limit(3).peek(n -> c[0]++).forEach(n -> {}); System.out.println(c[0]);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn stream_iterate_distinct_on_constant() {
    let out = run_main(
        "long c = java.util.stream.Stream.iterate(7, n -> 7).limit(4).distinct().count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn stream_iterate_reduce_sum() {
    let out = run_main(
        "int s = java.util.stream.Stream.iterate(1, n -> n + 1).limit(4).reduce(0, (a, b) -> a + b); System.out.println(s);",
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn stream_iterate_find_first() {
    let out = run_main(
        "int v = java.util.stream.Stream.iterate(5, n -> n + 2).filter(n -> n > 8).findFirst().get(); System.out.println(v);",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn stream_iterate_any_match() {
    let out = run_main(
        "boolean b = java.util.stream.Stream.iterate(1, n -> n + 1).limit(5).anyMatch(n -> n == 4); System.out.println(b);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn stream_iterate_all_match() {
    let out = run_main(
        "boolean b = java.util.stream.Stream.iterate(1, n -> n + 1).limit(3).allMatch(n -> n > 0); System.out.println(b);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn stream_iterate_none_match() {
    let out = run_main(
        "boolean b = java.util.stream.Stream.iterate(1, n -> n + 1).limit(3).noneMatch(n -> n < 0); System.out.println(b);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn stream_iterate_collect_to_list() {
    let out = run_main(
        "int sz = java.util.stream.Stream.iterate(0, n -> n + 1).limit(4).collect(java.util.stream.Collectors.toList()).size(); System.out.println(sz);",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn stream_iterate_predicate_stops_at_boundary() {
    let out = run_main(
        "long c = java.util.stream.Stream.iterate(10, n -> n > 5, n -> n - 2).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn stream_iterate_concat_with_finite() {
    let out = run_main(
        "long c = java.util.stream.Stream.concat(java.util.stream.Stream.iterate(1, n -> n + 1).limit(2), java.util.stream.Stream.of(99)).count(); System.out.println(c);",
    );
    assert_eq!(out, vec!["3"]);
}
