use crate::helpers::{run_in_main, run_main};

#[test]
fn optional_of_wraps_non_null_value_as_present() {
    let out = run_main(
        r#"java.util.Optional<String> opt = java.util.Optional.of("value"); System.out.println(opt.isPresent());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn optional_of_nullable_empty_when_source_is_null() {
    let out = run_main(
        "String s = null; java.util.Optional<String> opt = java.util.Optional.ofNullable(s); System.out.println(opt.isPresent());",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn optional_of_nullable_present_when_source_is_non_null() {
    let out = run_main(
        r#"String s = "live"; java.util.Optional<String> opt = java.util.Optional.ofNullable(s); System.out.println(opt.isPresent());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn optional_empty_is_not_present() {
    let out = run_main(
        "java.util.Optional<String> opt = java.util.Optional.empty(); System.out.println(opt.isPresent());",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn optional_map_transforms_present_value() {
    let out = run_main(
        r#"java.util.Optional<Integer> mapped = java.util.Optional.of(3).map(n -> n * 4); System.out.println(mapped.get());"#,
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn optional_map_on_empty_yields_empty_optional() {
    let out = run_main(
        "java.util.Optional<Integer> mapped = java.util.Optional.<Integer>empty().map(n -> n + 1); System.out.println(mapped.isPresent());",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn optional_map_chains_string_uppercase() {
    let out = run_main(
        r#"java.util.Optional<String> out = java.util.Optional.of("ok").map(s -> s.toUpperCase()); System.out.println(out.get());"#,
    );
    assert_eq!(out, vec!["OK"]);
}

#[test]
fn optional_flat_map_flattens_nested_optional() {
    let out = run_main(
        r#"java.util.Optional<String> out = java.util.Optional.of("7").flatMap(s -> java.util.Optional.of("len" + s.length())); System.out.println(out.get());"#,
    );
    assert_eq!(out, vec!["len1"]);
}

#[test]
fn optional_flat_map_on_empty_returns_empty() {
    let out = run_main(
        r#"java.util.Optional<String> out = java.util.Optional.<String>empty().flatMap(s -> java.util.Optional.of(s)); System.out.println(out.isPresent());"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn optional_flat_map_empty_inner_clears_present_value() {
    let out = run_main(
        r#"java.util.Optional<String> out = java.util.Optional.of("x").flatMap(s -> java.util.Optional.empty()); System.out.println(out.isPresent());"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn optional_filter_keeps_matching_present_value() {
    let out = run_main(
        "java.util.Optional<Integer> out = java.util.Optional.of(8).filter(n -> n > 5); System.out.println(out.get());",
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn optional_filter_discards_non_matching_present_value() {
    let out = run_main(
        "java.util.Optional<Integer> out = java.util.Optional.of(2).filter(n -> n > 5); System.out.println(out.isPresent());",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn optional_filter_on_empty_remains_empty() {
    let out = run_main(
        "java.util.Optional<Integer> out = java.util.Optional.<Integer>empty().filter(n -> n > 0); System.out.println(out.isPresent());",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn optional_or_else_supplies_default_for_empty() {
    let out = run_main(
        r#"java.util.Optional<String> opt = java.util.Optional.empty(); System.out.println(opt.orElse("fallback"));"#,
    );
    assert_eq!(out, vec!["fallback"]);
}

#[test]
fn optional_or_else_returns_value_when_present() {
    let out = run_main(
        r#"java.util.Optional<String> opt = java.util.Optional.of("real"); System.out.println(opt.orElse("fallback"));"#,
    );
    assert_eq!(out, vec!["real"]);
}

#[test]
fn optional_or_else_get_supplies_lazy_default_for_empty() {
    let out = run_main(
        r#"java.util.Optional<String> opt = java.util.Optional.empty(); System.out.println(opt.orElseGet(() -> "lazy"));"#,
    );
    assert_eq!(out, vec!["lazy"]);
}

#[test]
fn optional_or_else_get_skips_supplier_when_present() {
    let out = run_main(
        r#"java.util.Optional<String> opt = java.util.Optional.of("now"); System.out.println(opt.orElseGet(() -> "lazy"));"#,
    );
    assert_eq!(out, vec!["now"]);
}

#[test]
fn optional_or_else_throw_on_empty_throws_exception() {
    let out = run_in_main(
        "try { java.util.Optional.<String>empty().orElseThrow(() -> new RuntimeException(\"missing\")); System.out.println(\"ok\"); } catch (RuntimeException e) { System.out.println(e.getMessage()); }",
        "",
    );
    assert_eq!(out, vec!["missing"]);
}

#[test]
fn optional_or_else_throw_on_present_returns_value() {
    let out = run_main(
        r#"java.util.Optional<String> opt = java.util.Optional.of("safe"); System.out.println(opt.orElseThrow(() -> new RuntimeException("missing")));"#,
    );
    assert_eq!(out, vec!["safe"]);
}

#[test]
fn optional_if_present_runs_consumer_when_value_exists() {
    let out = run_main(
        r#"java.util.Optional<String> opt = java.util.Optional.of("hi"); opt.ifPresent(v -> System.out.println(v));"#,
    );
    assert_eq!(out, vec!["hi"]);
}

#[test]
fn optional_if_present_skips_consumer_when_empty() {
    let out = run_main(
        r#"java.util.Optional<String> opt = java.util.Optional.empty(); opt.ifPresent(v -> System.out.println(v)); System.out.println("done");"#,
    );
    assert_eq!(out, vec!["done"]);
}

#[test]
fn optional_is_present_false_for_empty() {
    let out = run_main(
        "java.util.Optional<Integer> opt = java.util.Optional.empty(); System.out.println(opt.isPresent());",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn optional_is_present_true_for_of_value() {
    let out = run_main(
        "java.util.Optional<Integer> opt = java.util.Optional.of(42); System.out.println(opt.isPresent());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn optional_stream_on_present_value_yields_singleton() {
    let out = run_main(
        "long count = java.util.Optional.of(9).stream().count(); System.out.println(count);",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn optional_stream_on_empty_value_yields_nothing() {
    let out = run_main(
        "long count = java.util.Optional.<Integer>empty().stream().count(); System.out.println(count);",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn optional_stream_map_in_pipeline_transforms_value() {
    let out = run_main(
        "int doubled = java.util.Optional.of(5).stream().mapToInt(n -> n * 2).sum(); System.out.println(doubled);",
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn optional_stream_for_each_prints_contained_value() {
    let out =
        run_main(r#"java.util.Optional.of("one").stream().forEach(v -> System.out.println(v));"#);
    assert_eq!(out, vec!["one"]);
}

#[test]
fn optional_get_on_present_returns_wrapped_value() {
    let out = run_main(
        r#"java.util.Optional<String> opt = java.util.Optional.of("data"); System.out.println(opt.get());"#,
    );
    assert_eq!(out, vec!["data"]);
}

#[test]
fn optional_of_nullable_from_method_returning_null() {
    let out = run_in_main(
        r#"java.util.Optional<Integer> opt = java.util.Optional.ofNullable(pickNull()); System.out.println(opt.isPresent());"#,
        r#"static Integer pickNull() { return null; }"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn optional_map_then_or_else_chains_transformation_and_default() {
    let out = run_main(
        r#"String out = java.util.Optional.of("3").map(s -> "n" + s).orElse("none"); System.out.println(out);"#,
    );
    assert_eq!(out, vec!["n3"]);
}

#[test]
fn optional_filter_then_map_keeps_and_transforms() {
    let out = run_main(
        "java.util.Optional<Integer> out = java.util.Optional.of(10).filter(n -> n > 5).map(n -> n + 1); System.out.println(out.get());",
    );
    assert_eq!(out, vec!["11"]);
}

#[test]
fn optional_flat_map_then_filter_composes_monadic_steps() {
    let out = run_main(
        r#"java.util.Optional<String> out = java.util.Optional.of("12").flatMap(s -> java.util.Optional.of(Integer.parseInt(s))).filter(n -> n > 10).map(n -> "big"); System.out.println(out.get());"#,
    );
    assert_eq!(out, vec!["big"]);
}

#[test]
fn optional_or_else_with_null_default_returns_null_string() {
    let out = run_main(
        r#"java.util.Optional<String> opt = java.util.Optional.empty(); String v = opt.orElse(null); System.out.println(v == null);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn optional_of_nullable_on_zero_integer_is_present() {
    let out = run_main(
        "Integer n = 0; java.util.Optional<Integer> opt = java.util.Optional.ofNullable(n); System.out.println(opt.isPresent());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn optional_map_to_string_length_from_optional_string() {
    let out = run_main(
        r#"java.util.Optional<Integer> len = java.util.Optional.of("abcd").map(String::length); System.out.println(len.get());"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn optional_empty_stream_collects_to_empty_list() {
    let out = run_main(
        "java.util.List<Integer> list = java.util.Optional.<Integer>empty().stream().toList(); System.out.println(list.size());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn optional_present_stream_collects_to_singleton_list() {
    let out = run_main(
        "java.util.List<Integer> list = java.util.Optional.of(7).stream().toList(); System.out.println(list.size()); System.out.println(list.get(0));",
    );
    assert_eq!(out, vec!["1", "7"]);
}

#[test]
fn optional_filter_rejects_nullable_empty_string() {
    let out = run_main(
        r#"String s = ""; java.util.Optional<String> out = java.util.Optional.ofNullable(s).filter(v -> v.length() > 0); System.out.println(out.isPresent());"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn optional_or_else_get_with_side_effect_only_when_empty() {
    let out = run_main(
        r#"int[] flag = new int[]{0}; String v = java.util.Optional.<String>empty().orElseGet(() -> { flag[0] = 1; return "x"; }); System.out.println(flag[0]); System.out.println(v);"#,
    );
    assert_eq!(out, vec!["1", "x"]);
}

#[test]
fn optional_if_present_or_else_runs_alternate_when_empty() {
    let out = run_main(
        r#"java.util.Optional.<String>empty().ifPresentOrElse(v -> System.out.println(v), () -> System.out.println("none"));"#,
    );
    assert_eq!(out, vec!["none"]);
}

#[test]
fn optional_if_present_or_else_runs_present_action_when_value_exists() {
    let out = run_main(
        r#"java.util.Optional.of("yes").ifPresentOrElse(v -> System.out.println(v), () -> System.out.println("none"));"#,
    );
    assert_eq!(out, vec!["yes"]);
}

#[test]
fn optional_nested_of_nullable_outer_inner_chain() {
    let out = run_main(
        r#"String s = "inner"; java.util.Optional<String> opt = java.util.Optional.ofNullable(java.util.Optional.ofNullable(s).orElse(null)); System.out.println(opt.orElse("missing"));"#,
    );
    assert_eq!(out, vec!["inner"]);
}

#[test]
fn optional_stream_filter_in_pipeline_keeps_matches() {
    let out = run_main(
        "long count = java.util.Optional.of(6).stream().filter(n -> n % 2 == 0).count(); System.out.println(count);",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn optional_equals_on_same_present_values() {
    let out = run_main(
        r#"java.util.Optional<String> a = java.util.Optional.of("z"); java.util.Optional<String> b = java.util.Optional.of("z"); System.out.println(a.equals(b));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn optional_to_string_on_empty_contains_empty_token() {
    let out = run_main(
        r#"String s = java.util.Optional.empty().toString(); System.out.println(s.contains("empty"));"#,
    );
    assert_eq!(out, vec!["true"]);
}
