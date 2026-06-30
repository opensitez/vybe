use crate::helpers::run_main;

#[test]
fn unary_operator_identity_returns_same_integer() {
    let out = run_main(
        "java.util.function.UnaryOperator<Integer> id = java.util.function.UnaryOperator.identity(); System.out.println(id.apply(7));",
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn unary_operator_identity_on_string() {
    let out = run_main(
        r#"java.util.function.UnaryOperator<String> id = java.util.function.UnaryOperator.identity(); System.out.println(id.apply("keep"));"#,
    );
    assert_eq!(out, vec!["keep"]);
}

#[test]
fn unary_operator_compose_applies_functions_in_order() {
    let out = run_main(
        "java.util.function.UnaryOperator<Integer> inc = n -> n + 1; java.util.function.UnaryOperator<Integer> doubleIt = n -> n * 2; System.out.println(inc.compose(doubleIt).apply(3));",
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn unary_operator_and_then_chains_transformations() {
    let out = run_main(
        "java.util.function.UnaryOperator<Integer> inc = n -> n + 1; java.util.function.UnaryOperator<Integer> doubleIt = n -> n * 2; System.out.println(inc.andThen(doubleIt).apply(3));",
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn binary_operator_max_by_natural_order() {
    let out = run_main(
        "java.util.function.BinaryOperator<Integer> max = java.util.function.BinaryOperator.maxBy(Integer::compareTo); System.out.println(max.apply(3, 9));",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn binary_operator_min_by_natural_order() {
    let out = run_main(
        "java.util.function.BinaryOperator<Integer> min = java.util.function.BinaryOperator.minBy(Integer::compareTo); System.out.println(min.apply(3, 9));",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn binary_operator_max_by_string_length() {
    let out = run_main(
        "java.util.function.BinaryOperator<String> max = java.util.function.BinaryOperator.maxBy((a, b) -> a.length() - b.length()); System.out.println(max.apply(\"aa\", \"bbb\"));",
    );
    assert_eq!(out, vec!["bbb"]);
}

#[test]
fn binary_operator_min_by_string_length() {
    let out = run_main(
        "java.util.function.BinaryOperator<String> min = java.util.function.BinaryOperator.minBy((a, b) -> a.length() - b.length()); System.out.println(min.apply(\"aa\", \"bbb\"));",
    );
    assert_eq!(out, vec!["aa"]);
}

#[test]
fn int_binary_operator_sum_in_reduce() {
    let out = run_main(
        "int s = java.util.stream.IntStream.of(1, 2, 3).reduce(0, (a, b) -> a + b); System.out.println(s);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn int_binary_operator_max_via_method_reference() {
    let out = run_main(
        "int v = java.util.stream.IntStream.of(4, 9, 2).reduce(Integer::max).getAsInt(); System.out.println(v);",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn long_binary_operator_product() {
    let out = run_main(
        "long p = java.util.stream.LongStream.of(2L, 3L, 4L).reduce(1L, (a, b) -> a * b); System.out.println(p);",
    );
    assert_eq!(out, vec!["24"]);
}

#[test]
fn double_binary_operator_sum() {
    let out = run_main(
        "double s = java.util.stream.DoubleStream.of(1.5, 2.5).reduce(0.0, (a, b) -> a + b); System.out.println((int) s);",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn binary_operator_custom_string_concat() {
    let out = run_main(
        "java.util.function.BinaryOperator<String> join = (a, b) -> a + \":\" + b; System.out.println(join.apply(\"x\", \"y\"));",
    );
    assert_eq!(out, vec!["x:y"]);
}

#[test]
fn unary_operator_on_list_replace_all() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); list.replaceAll(n -> n * 10); System.out.println(list.get(1));",
    );
    assert_eq!(out, vec!["20"]);
}

#[test]
fn unary_operator_map_entry_transform() {
    let out = run_main(
        r#"java.util.HashMap<String, Integer> map = new java.util.HashMap<String, Integer>(); map.put("a", 1); map.replaceAll((k, v) -> v + 1); System.out.println(map.get("a"));"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn binary_operator_min_by_after_map_in_stream() {
    let out = run_main(
        "java.util.Optional<Integer> m = java.util.Arrays.asList(3, 9, 5).stream().collect(java.util.stream.Collectors.minBy(Integer::compareTo)); System.out.println(m.get());",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn binary_operator_max_by_after_filter() {
    let out = run_main(
        "java.util.Optional<Integer> m = java.util.Arrays.asList(1, 8, 3, 6).stream().filter(n -> n < 7).collect(java.util.stream.Collectors.maxBy(Integer::compareTo)); System.out.println(m.get());",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn unary_operator_and_then_with_string_uppercase() {
    let out = run_main(
        r#"java.util.function.UnaryOperator<String> trim = String::strip; java.util.function.UnaryOperator<String> upper = String::toUpperCase; System.out.println(trim.andThen(upper).apply("  hi  "));"#,
    );
    assert_eq!(out, vec!["HI"]);
}

#[test]
fn unary_operator_compose_with_string_length_tag() {
    let out = run_main(
        r#"java.util.function.UnaryOperator<String> tag = s -> "[" + s + "]"; java.util.function.UnaryOperator<String> upper = String::toUpperCase; System.out.println(tag.compose(upper).apply("ok"));"#,
    );
    assert_eq!(out, vec!["[OK]"]);
}

#[test]
fn int_binary_operator_bitwise_or() {
    let out = run_main(
        "int v = java.util.stream.IntStream.of(4, 2).reduce((a, b) -> a | b).getAsInt(); System.out.println(v);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn long_binary_operator_bitwise_and() {
    let out = run_main(
        "long v = java.util.stream.LongStream.of(7L, 3L).reduce((a, b) -> a & b).getAsLong(); System.out.println(v);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn double_binary_operator_max_of_pair() {
    let out = run_main(
        "double v = java.util.stream.DoubleStream.of(2.5, 4.5).reduce(Math::max).getAsDouble(); System.out.println((int) v);",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn binary_operator_min_by_descending_comparator() {
    let out = run_main(
        "java.util.function.BinaryOperator<Integer> min = java.util.function.BinaryOperator.minBy((a, b) -> b - a); System.out.println(min.apply(3, 9));",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn binary_operator_max_by_descending_comparator() {
    let out = run_main(
        "java.util.function.BinaryOperator<Integer> max = java.util.function.BinaryOperator.maxBy((a, b) -> b - a); System.out.println(max.apply(3, 9));",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn unary_operator_identity_compose_is_noop() {
    let out = run_main(
        "java.util.function.UnaryOperator<Integer> inc = n -> n + 1; System.out.println(inc.compose(java.util.function.UnaryOperator.identity()).apply(4));",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn unary_operator_identity_and_then_is_noop() {
    let out = run_main(
        "java.util.function.UnaryOperator<Integer> inc = n -> n + 1; System.out.println(inc.andThen(java.util.function.UnaryOperator.identity()).apply(4));",
    );
    assert_eq!(out, vec!["5"]);
}
