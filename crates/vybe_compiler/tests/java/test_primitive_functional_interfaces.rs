use crate::helpers::run_main;

#[test]
fn int_predicate_test_even() {
    let out = run_main(
        "java.util.function.IntPredicate even = n -> n % 2 == 0; System.out.println(even.test(4)); System.out.println(even.test(5));",
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn int_predicate_and_both_true() {
    let out = run_main(
        "java.util.function.IntPredicate p = n -> n > 0; System.out.println(p.and(n -> n < 10).test(5));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn int_predicate_or_either_true() {
    let out = run_main(
        "java.util.function.IntPredicate p = n -> n < 0; System.out.println(p.or(n -> n == 7).test(7));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn int_predicate_negate() {
    let out = run_main(
        "java.util.function.IntPredicate p = n -> n > 5; System.out.println(p.negate().test(3));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn int_consumer_accepts_and_prints() {
    let out =
        run_main("java.util.function.IntConsumer c = n -> System.out.println(n); c.accept(9);");
    assert_eq!(out, vec!["9"]);
}

#[test]
fn int_consumer_and_then_sequence() {
    let out = run_main(
        "StringBuilder sb = new StringBuilder(); java.util.function.IntConsumer c = n -> sb.append(n); c.andThen(n -> sb.append(\"!\")).accept(3); System.out.println(sb.toString());",
    );
    assert_eq!(out, vec!["3!"]);
}

#[test]
fn int_function_apply_square() {
    let out = run_main(
        "java.util.function.IntFunction<Integer> sq = n -> n * n; System.out.println(sq.apply(4));",
    );
    assert_eq!(out, vec!["16"]);
}

#[test]
fn int_function_and_then_to_string() {
    let out = run_main(
        "java.util.function.IntFunction<String> tag = n -> \"v\" + n; System.out.println(tag.andThen(String::toUpperCase).apply(2));",
    );
    assert_eq!(out, vec!["V2"]);
}

#[test]
fn int_to_long_function_widens() {
    let out = run_main(
        "java.util.function.IntToLongFunction widen = n -> (long) n * 10L; System.out.println(widen.applyAsLong(3));",
    );
    assert_eq!(out, vec!["30"]);
}

#[test]
fn int_to_double_function_halves() {
    let out = run_main(
        "java.util.function.IntToDoubleFunction half = n -> n / 2.0; System.out.println((int) half.applyAsDouble(5));",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn int_unary_operator_increment() {
    let out = run_main(
        "java.util.function.IntUnaryOperator inc = n -> n + 1; System.out.println(inc.applyAsInt(8));",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn int_unary_operator_and_then_double() {
    let out = run_main(
        "java.util.function.IntUnaryOperator f = n -> n + 1; System.out.println(f.andThen(n -> n * 2).applyAsInt(3));",
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn int_unary_operator_compose() {
    let out = run_main(
        "java.util.function.IntUnaryOperator f = n -> n * 2; System.out.println(f.compose(n -> n + 1).applyAsInt(3));",
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn int_supplier_constant() {
    let out =
        run_main("java.util.function.IntSupplier s = () -> 42; System.out.println(s.getAsInt());");
    assert_eq!(out, vec!["42"]);
}

#[test]
fn long_predicate_test_positive() {
    let out = run_main(
        "java.util.function.LongPredicate p = n -> n > 0L; System.out.println(p.test(1L));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn long_predicate_and() {
    let out = run_main(
        "java.util.function.LongPredicate p = n -> n > 0L; System.out.println(p.and(n -> n < 5L).test(3L));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn long_predicate_or() {
    let out = run_main(
        "java.util.function.LongPredicate p = n -> n < 0L; System.out.println(p.or(n -> n == 10L).test(10L));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn long_predicate_negate() {
    let out = run_main(
        "java.util.function.LongPredicate p = n -> n > 5L; System.out.println(p.negate().test(3L));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn long_consumer_and_then() {
    let out = run_main(
        "StringBuilder sb = new StringBuilder(); java.util.function.LongConsumer c = n -> sb.append(n); c.andThen(n -> sb.append(\"L\")).accept(5L); System.out.println(sb.toString());",
    );
    assert_eq!(out, vec!["5L"]);
}

#[test]
fn long_function_apply() {
    let out = run_main(
        "java.util.function.LongFunction<String> f = n -> \"L\" + n; System.out.println(f.apply(7L));",
    );
    assert_eq!(out, vec!["L7"]);
}

#[test]
fn long_to_int_function_truncates() {
    let out = run_main(
        "java.util.function.LongToIntFunction f = n -> (int) n; System.out.println(f.applyAsInt(9L));",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn long_to_double_function() {
    let out = run_main(
        "java.util.function.LongToDoubleFunction f = n -> (double) n; System.out.println((int) f.applyAsDouble(4L));",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn long_unary_operator_negate() {
    let out = run_main(
        "java.util.function.LongUnaryOperator f = n -> -n; System.out.println(f.applyAsLong(5L));",
    );
    assert_eq!(out, vec!["-5"]);
}

#[test]
fn long_supplier_reads_capture() {
    let out = run_main(
        "long base = 3L; java.util.function.LongSupplier s = () -> base + 1L; System.out.println(s.getAsLong());",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn double_predicate_test_finite() {
    let out = run_main(
        "java.util.function.DoublePredicate p = n -> !Double.isNaN(n); System.out.println(p.test(1.0));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn double_predicate_and() {
    let out = run_main(
        "java.util.function.DoublePredicate p = n -> n > 0.0; System.out.println(p.and(n -> n < 10.0).test(5.0));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn double_predicate_or() {
    let out = run_main(
        "java.util.function.DoublePredicate p = n -> n < 0.0; System.out.println(p.or(n -> n == 2.5).test(2.5));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn double_predicate_negate() {
    let out = run_main(
        "java.util.function.DoublePredicate p = n -> n > 1.0; System.out.println(p.negate().test(0.5));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn double_consumer_and_then() {
    let out = run_main(
        "StringBuilder sb = new StringBuilder(); java.util.function.DoubleConsumer c = n -> sb.append((int) n); c.andThen(n -> sb.append(\"D\")).accept(3.0); System.out.println(sb.toString());",
    );
    assert_eq!(out, vec!["3D"]);
}

#[test]
fn double_function_apply() {
    let out = run_main(
        "java.util.function.DoubleFunction<String> f = n -> String.valueOf((int) n); System.out.println(f.apply(4.9));",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn double_to_int_function() {
    let out = run_main(
        "java.util.function.DoubleToIntFunction f = n -> (int) n; System.out.println(f.applyAsInt(7.9));",
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn double_to_long_function() {
    let out = run_main(
        "java.util.function.DoubleToLongFunction f = n -> (long) n; System.out.println(f.applyAsLong(8.9));",
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn double_unary_operator_abs() {
    let out = run_main(
        "java.util.function.DoubleUnaryOperator f = Math::abs; System.out.println((int) f.applyAsDouble(-3.0));",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn double_supplier_pi() {
    let out = run_main(
        "java.util.function.DoubleSupplier s = () -> 3.14; System.out.println((int) s.getAsDouble());",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn obj_int_consumer_bi_accepts() {
    let out = run_main(
        "java.util.function.ObjIntConsumer<String> c = (s, n) -> System.out.println(s + n); c.accept(\"x\", 2);",
    );
    assert_eq!(out, vec!["x2"]);
}

#[test]
fn obj_long_consumer_bi_accepts() {
    let out = run_main(
        "java.util.function.ObjLongConsumer<String> c = (s, n) -> System.out.println(s.length() + (int) n); c.accept(\"ab\", 1L);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn obj_double_consumer_bi_accepts() {
    let out = run_main(
        "java.util.function.ObjDoubleConsumer<String> c = (s, n) -> System.out.println(s + (int) n); c.accept(\"v\", 3.0);",
    );
    assert_eq!(out, vec!["v3"]);
}

#[test]
fn to_int_function_string_length() {
    let out = run_main(
        "java.util.function.ToIntFunction<String> len = String::length; System.out.println(len.applyAsInt(\"java\"));",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn to_long_function_parse() {
    let out = run_main(
        "java.util.function.ToLongFunction<String> parse = Long::parseLong; System.out.println(parse.applyAsLong(\"42\"));",
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn to_double_function_parse() {
    let out = run_main(
        "java.util.function.ToDoubleFunction<String> parse = Double::parseDouble; System.out.println((int) parse.applyAsDouble(\"3.5\"));",
    );
    assert_eq!(out, vec!["3"]);
}
