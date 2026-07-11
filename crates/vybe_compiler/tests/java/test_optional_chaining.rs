use crate::helpers::run_main;

#[test]
fn optional_is_empty_true_for_empty() {
    let out = run_main("System.out.println(java.util.Optional.empty().isEmpty());");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn optional_is_empty_false_for_present() {
    let out = run_main("System.out.println(java.util.Optional.of(1).isEmpty());");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn optional_hash_code_present() {
    let out = run_main(
        "System.out.println(java.util.Optional.of(\"a\").hashCode() == java.util.Optional.of(\"a\").hashCode());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn optional_hash_code_empty_constant() {
    let out = run_main(
        "int h1 = java.util.Optional.empty().hashCode(); int h2 = java.util.Optional.empty().hashCode(); System.out.println(h1 == h2);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn optional_equals_empty_to_empty() {
    let out = run_main(
        "System.out.println(java.util.Optional.empty().equals(java.util.Optional.empty()));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn optional_equals_present_to_different() {
    let out =
        run_main("System.out.println(java.util.Optional.of(1).equals(java.util.Optional.of(2)));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn optional_or_returns_other_when_empty() {
    let out = run_main(
        "java.util.Optional<String> a = java.util.Optional.empty(); java.util.Optional<String> b = java.util.Optional.of(\"fallback\"); System.out.println(a.or(b).get());",
    );
    assert_eq!(out, vec!["fallback"]);
}

#[test]
fn optional_or_keeps_self_when_present() {
    let out = run_main(
        "java.util.Optional<String> a = java.util.Optional.of(\"keep\"); java.util.Optional<String> b = java.util.Optional.of(\"other\"); System.out.println(a.or(b).get());",
    );
    assert_eq!(out, vec!["keep"]);
}

#[test]
fn optional_or_empty_other_stays_empty() {
    let out = run_main(
        "java.util.Optional<String> a = java.util.Optional.empty(); java.util.Optional<String> b = java.util.Optional.empty(); System.out.println(a.or(b).isEmpty());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn optional_or_else_throw_no_arg_on_empty() {
    let out = run_main(
        "try { java.util.Optional.empty().orElseThrow(); System.out.println(\"ok\"); } catch (java.util.NoSuchElementException e) { System.out.println(\"missing\"); }",
    );
    assert_eq!(out, vec!["missing"]);
}

#[test]
fn optional_or_else_throw_no_arg_on_present() {
    let out = run_main("System.out.println(java.util.Optional.of(7).orElseThrow());");
    assert_eq!(out, vec!["7"]);
}

#[test]
fn optional_if_present_or_else_empty_branch() {
    let out = run_main(
        "java.util.Optional.<String>empty().ifPresentOrElse(v -> System.out.println(\"v\"), () -> System.out.println(\"empty\"));",
    );
    assert_eq!(out, vec!["empty"]);
}

#[test]
fn optional_stream_reduce_sums_present() {
    let out = run_main(
        "int s = java.util.Optional.of(5).stream().reduce(0, (a, b) -> a + b); System.out.println(s);",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn optional_stream_reduce_empty_is_identity() {
    let out = run_main(
        "int s = java.util.Optional.<Integer>empty().stream().reduce(9, (a, b) -> a + b); System.out.println(s);",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn optional_map_null_result_becomes_empty() {
    let out = run_main(
        "java.util.Optional<String> out = java.util.Optional.of(\"x\").map(s -> null); System.out.println(out.isEmpty());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn optional_flat_map_to_empty_optional() {
    let out = run_main(
        "java.util.Optional<Integer> out = java.util.Optional.of(1).flatMap(n -> java.util.Optional.empty()); System.out.println(out.isEmpty());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn optional_filter_then_or_else_chain() {
    let out = run_main(
        "String v = java.util.Optional.of(4).filter(n -> n > 5).map(n -> \"\" + n).orElse(\"none\"); System.out.println(v);",
    );
    assert_eq!(out, vec!["none"]);
}

#[test]
fn optional_nested_or_picks_first_present() {
    let out = run_main(
        "java.util.Optional<String> v = java.util.Optional.<String>empty().or(() -> java.util.Optional.of(\"inner\")); System.out.println(v.get());",
    );
    assert_eq!(out, vec!["inner"]);
}

#[test]
fn optional_to_string_present_contains_value() {
    let out = run_main(
        "String s = java.util.Optional.of(42).toString(); System.out.println(s.contains(\"42\"));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn optional_equals_present_same_type_different_value() {
    let out =
        run_main("System.out.println(java.util.Optional.of(1).equals(java.util.Optional.of(1L)));");
    assert_eq!(out, vec!["false"]);
}
