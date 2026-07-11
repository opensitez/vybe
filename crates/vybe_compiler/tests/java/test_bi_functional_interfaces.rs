use crate::helpers::run_main;

#[test]
fn bi_predicate_test_both_args() {
    let out = run_main(
        "java.util.function.BiPredicate<String, Integer> p = (s, n) -> s.length() == n; System.out.println(p.test(\"ab\", 2));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn bi_predicate_and() {
    let out = run_main(
        "java.util.function.BiPredicate<Integer, Integer> p = (a, b) -> a > 0; System.out.println(p.and((a, b) -> b > 0).test(2, 3));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn bi_predicate_or() {
    let out = run_main(
        "java.util.function.BiPredicate<Integer, Integer> p = (a, b) -> a < 0; System.out.println(p.or((a, b) -> b == 5).test(1, 5));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn bi_predicate_negate() {
    let out = run_main(
        "java.util.function.BiPredicate<Integer, Integer> p = (a, b) -> a.equals(b); System.out.println(p.negate().test(1, 2));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn bi_predicate_is_equal() {
    let out = run_main(
        "java.util.function.BiPredicate<String, String> p = java.util.function.BiPredicate.isEqual(\"x\"); System.out.println(p.test(\"x\", \"x\"));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn bi_function_apply_concat() {
    let out = run_main(
        "java.util.function.BiFunction<String, String, String> f = (a, b) -> a + b; System.out.println(f.apply(\"a\", \"b\"));",
    );
    assert_eq!(out, vec!["ab"]);
}

#[test]
fn bi_function_and_then_uppercase() {
    let out = run_main(
        "java.util.function.BiFunction<String, String, String> f = (a, b) -> a + b; System.out.println(f.andThen(String::toUpperCase).apply(\"a\", \"b\"));",
    );
    assert_eq!(out, vec!["AB"]);
}

#[test]
fn bi_function_apply_then_identity_and_then() {
    let out = run_main(
        "java.util.function.BiFunction<Integer, Integer, Integer> add = (a, b) -> a + b; System.out.println(add.andThen(n -> n).andThen(n -> n * 2).apply(2, 3));",
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn bi_consumer_accept_and_print() {
    let out = run_main(
        "java.util.function.BiConsumer<String, Integer> c = (s, n) -> System.out.println(s + n); c.accept(\"n\", 7);",
    );
    assert_eq!(out, vec!["n7"]);
}

#[test]
fn bi_consumer_and_then() {
    let out = run_main(
        "StringBuilder sb = new StringBuilder(); java.util.function.BiConsumer<String, String> c = (a, b) -> sb.append(a); c.andThen((a, b) -> sb.append(b)).accept(\"x\", \"y\"); System.out.println(sb.toString());",
    );
    assert_eq!(out, vec!["xy"]);
}

#[test]
fn binary_operator_max() {
    let out = run_main(
        "java.util.function.BinaryOperator<Integer> max = (a, b) -> a > b ? a : b; System.out.println(max.apply(3, 9));",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn binary_operator_min_by_comparator() {
    let out = run_main(
        "java.util.function.BinaryOperator<String> min = java.util.function.BinaryOperator.minBy(String::compareTo); System.out.println(min.apply(\"b\", \"a\"));",
    );
    assert_eq!(out, vec!["a"]);
}

#[test]
fn binary_operator_max_by_comparator() {
    let out = run_main(
        "java.util.function.BinaryOperator<String> max = java.util.function.BinaryOperator.maxBy(String::compareTo); System.out.println(max.apply(\"b\", \"a\"));",
    );
    assert_eq!(out, vec!["b"]);
}

#[test]
fn bi_predicate_and_short_circuit() {
    let out = run_main(
        "java.util.function.BiPredicate<Integer, Integer> fail = (a, b) -> false; java.util.function.BiPredicate<Integer, Integer> boom = (a, b) -> { throw new RuntimeException(); }; System.out.println(fail.and(boom).test(1, 2));",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn bi_predicate_or_short_circuit() {
    let out = run_main(
        "java.util.function.BiPredicate<Integer, Integer> pass = (a, b) -> true; java.util.function.BiPredicate<Integer, Integer> boom = (a, b) -> { throw new RuntimeException(); }; System.out.println(pass.or(boom).test(1, 2));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn bi_function_and_then_parse_sum() {
    let out = run_main(
        "java.util.function.BiFunction<String, String, Integer> concat = (a, b) -> a + b; System.out.println(concat.andThen(Integer::parseInt).apply(\"1\", \"2\"));",
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn bi_consumer_and_then_three() {
    let out = run_main(
        "int[] c = {0}; java.util.function.BiConsumer<Integer, Integer> bc = (a, b) -> c[0] += a; bc.andThen((a, b) -> c[0] += b).accept(2, 3); System.out.println(c[0]);",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn bi_predicate_negate_is_equal() {
    let out = run_main(
        "java.util.function.BiPredicate<String, String> p = java.util.function.BiPredicate.isEqual(\"a\").negate(); System.out.println(p.test(\"a\", \"b\"));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn binary_operator_add_strings() {
    let out = run_main(
        "java.util.function.BinaryOperator<String> add = (a, b) -> a + b; System.out.println(add.apply(\"1\", \"23\"));",
    );
    assert_eq!(out, vec!["123"]);
}

#[test]
fn bi_function_identity_style_map() {
    let out = run_main(
        "java.util.function.BiFunction<Integer, Integer, Integer> mul = (a, b) -> a * b; System.out.println(mul.apply(3, 4));",
    );
    assert_eq!(out, vec!["12"]);
}
