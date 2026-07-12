/// Stream.generate — infinite supplier streams with short-circuit ops.
use crate::helpers::run_main;

#[test]
fn stream_generate_with_limit_produces_constant_values() {
    let out = run_main(
        r#"long count = java.util.stream.Stream.generate(() -> "x").limit(3).count(); System.out.println(count);"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn stream_generate_supplier_increments_counter_via_mutable_holder() {
    let out = run_main(
        r#"int[] box = {0}; java.util.List<Integer> list = java.util.stream.Stream.generate(() -> ++box[0]).limit(4).toList(); System.out.println(list.get(0)); System.out.println(list.get(3));"#,
    );
    assert_eq!(out, vec!["1", "4"]);
}

#[test]
fn stream_generate_concatenated_with_finite_stream() {
    let out = run_main(
        r#"java.util.List<String> list = java.util.stream.Stream.concat(java.util.stream.Stream.of("start"), java.util.stream.Stream.generate(() -> "tick").limit(2)).toList(); System.out.println(list.get(0)); System.out.println(list.get(2));"#,
    );
    assert_eq!(out, vec!["start", "tick"]);
}

#[test]
fn stream_generate_find_first_returns_initial_element() {
    let out = run_main(
        r#"java.util.Optional<String> first = java.util.stream.Stream.generate(() -> "loop").limit(5).findFirst(); System.out.println(first.get());"#,
    );
    assert_eq!(out, vec!["loop"]);
}

#[test]
fn stream_generate_filtered_by_predicate_before_limit() {
    let out = run_main(
        r#"int[] n = {0}; long count = java.util.stream.Stream.generate(() -> ++n[0]).filter(v -> v % 2 == 0).limit(2).count(); System.out.println(count);"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn stream_generate_collect_joins_constant_strings() {
    let out = run_main(
        r#"String joined = java.util.stream.Stream.generate(() -> "a").limit(3).collect(java.util.stream.Collectors.joining()); System.out.println(joined);"#,
    );
    assert_eq!(out, vec!["aaa"]);
}
