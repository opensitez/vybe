use crate::helpers::run_main;

#[test]
fn predicate_and_both_must_pass() {
    let out = run_main(
        "java.util.function.Predicate<Integer> gt2 = n -> n > 2; java.util.function.Predicate<Integer> lt5 = n -> n < 5; System.out.println(gt2.and(lt5).test(3)); System.out.println(gt2.and(lt5).test(6));",
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn predicate_and_short_circuits_on_first_false() {
    let out = run_main(
        "java.util.function.Predicate<Integer> fail = n -> false; java.util.function.Predicate<Integer> never = n -> { throw new RuntimeException(\"x\"); }; System.out.println(fail.and(never).test(1));",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn predicate_or_either_can_pass() {
    let out = run_main(
        "java.util.function.Predicate<Integer> lt2 = n -> n < 2; java.util.function.Predicate<Integer> gt5 = n -> n > 5; System.out.println(lt2.or(gt5).test(1)); System.out.println(lt2.or(gt5).test(7)); System.out.println(lt2.or(gt5).test(3));",
    );
    assert_eq!(out, vec!["true", "true", "false"]);
}

#[test]
fn predicate_or_short_circuits_on_first_true() {
    let out = run_main(
        "java.util.function.Predicate<Integer> pass = n -> true; java.util.function.Predicate<Integer> never = n -> { throw new RuntimeException(\"x\"); }; System.out.println(pass.or(never).test(1));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn predicate_negate_inverts_result() {
    let out = run_main(
        "java.util.function.Predicate<Integer> even = n -> n % 2 == 0; System.out.println(even.negate().test(3)); System.out.println(even.negate().test(4));",
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn predicate_is_equal_compares_to_target() {
    let out = run_main(
        "java.util.function.Predicate<String> isJava = java.util.function.Predicate.isEqual(\"java\"); System.out.println(isJava.test(\"java\")); System.out.println(isJava.test(\"go\"));",
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn predicate_and_chains_three_conditions() {
    let out = run_main(
        "java.util.function.Predicate<Integer> p = n -> n > 0; System.out.println(p.and(n -> n < 10).and(n -> n % 2 == 0).test(6));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn predicate_or_chains_three_alternatives() {
    let out = run_main(
        "java.util.function.Predicate<Integer> p = n -> n == 1; System.out.println(p.or(n -> n == 2).or(n -> n == 3).test(3));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn predicate_negate_of_is_equal() {
    let out = run_main(
        "java.util.function.Predicate<String> notA = java.util.function.Predicate.isEqual(\"a\").negate(); System.out.println(notA.test(\"b\"));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn function_compose_applies_outer_after_inner() {
    let out = run_main(
        "java.util.function.Function<Integer, Integer> doubleIt = n -> n * 2; java.util.function.Function<Integer, Integer> addOne = n -> n + 1; System.out.println(addOne.compose(doubleIt).apply(3));",
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn function_and_then_applies_inner_then_outer() {
    let out = run_main(
        "java.util.function.Function<Integer, Integer> doubleIt = n -> n * 2; java.util.function.Function<Integer, Integer> addOne = n -> n + 1; System.out.println(doubleIt.andThen(addOne).apply(3));",
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn function_compose_with_string_conversion() {
    let out = run_main(
        "java.util.function.Function<String, Integer> len = String::length; java.util.function.Function<Integer, String> tag = n -> \"n\" + n; System.out.println(tag.compose(len).apply(\"abc\"));",
    );
    assert_eq!(out, vec!["n3"]);
}

#[test]
fn function_and_then_parses_after_double() {
    let out = run_main(
        "java.util.function.Function<String, String> upper = String::toUpperCase; java.util.function.Function<String, Integer> len = String::length; System.out.println(upper.andThen(len).apply(\"hi\"));",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn function_identity_returns_argument() {
    let out = run_main(
        "java.util.function.Function<String, String> id = java.util.function.Function.identity(); System.out.println(id.apply(\"same\"));",
    );
    assert_eq!(out, vec!["same"]);
}

#[test]
fn consumer_and_then_runs_both_in_order() {
    let out = run_main(
        "StringBuilder sb = new StringBuilder(); java.util.function.Consumer<String> c1 = s -> sb.append(s); java.util.function.Consumer<String> c2 = s -> sb.append(s.toUpperCase()); c1.andThen(c2).accept(\"a\"); System.out.println(sb.toString());",
    );
    assert_eq!(out, vec!["aA"]);
}

#[test]
fn consumer_and_then_three_stages() {
    let out = run_main(
        "StringBuilder sb = new StringBuilder(); java.util.function.Consumer<Integer> c = n -> sb.append(n); c.andThen(n -> sb.append(\"-\")).andThen(n -> sb.append(\"!\")).accept(1); System.out.println(sb.toString());",
    );
    assert_eq!(out, vec!["1-!"]);
}

#[test]
fn predicate_and_with_method_reference() {
    let out = run_main(
        "java.util.function.Predicate<String> nonEmpty = s -> !s.isEmpty(); System.out.println(nonEmpty.and(String::isBlank).negate().test(\"x\"));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn function_compose_chain_three_functions() {
    let out = run_main(
        "java.util.function.Function<Integer, Integer> f = n -> n + 1; java.util.function.Function<Integer, Integer> g = n -> n * 2; java.util.function.Function<Integer, Integer> h = n -> n - 1; System.out.println(f.compose(g.compose(h)).apply(5));",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn function_and_then_chain_three_functions() {
    let out = run_main(
        "java.util.function.Function<Integer, Integer> f = n -> n + 1; java.util.function.Function<Integer, Integer> g = n -> n * 2; System.out.println(f.andThen(g).andThen(n -> n - 1).apply(3));",
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn consumer_and_then_prints_sequence() {
    let out = run_main(
        "java.util.function.Consumer<String> first = s -> System.out.println(s); java.util.function.Consumer<String> second = s -> System.out.println(s.length()); first.andThen(second).accept(\"ab\");",
    );
    assert_eq!(out, vec!["ab", "2"]);
}

#[test]
fn predicate_or_negated_and() {
    let out = run_main(
        "java.util.function.Predicate<Integer> even = n -> n % 2 == 0; System.out.println(even.negate().or(n -> n > 10).test(11));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn predicate_and_negated_or() {
    let out = run_main(
        "java.util.function.Predicate<Integer> small = n -> n < 5; System.out.println(small.and(n -> n > 0).or(n -> n == 99).test(99));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn function_compose_identity_is_noop() {
    let out = run_main(
        "java.util.function.Function<Integer, Integer> inc = n -> n + 1; System.out.println(inc.compose(java.util.function.Function.identity()).apply(4));",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn function_and_then_identity_is_noop() {
    let out = run_main(
        "java.util.function.Function<Integer, Integer> inc = n -> n + 1; System.out.println(inc.andThen(java.util.function.Function.identity()).apply(4));",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn consumer_and_then_with_side_effect_counter() {
    let out = run_main(
        "int[] c = {0}; java.util.function.Consumer<Integer> bump = n -> c[0]++; bump.andThen(n -> c[0] += 10).accept(0); System.out.println(c[0]);",
    );
    assert_eq!(out, vec!["11"]);
}

