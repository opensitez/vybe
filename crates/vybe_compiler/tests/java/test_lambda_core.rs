use crate::helpers::run_main;

#[test]
fn lambda_as_int_unary_operator_doubles_argument() {
    let out = run_main(
        "java.util.function.IntUnaryOperator doubler = x -> x * 2; System.out.println(doubler.applyAsInt(6));",
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn lambda_as_predicate_tests_even_numbers() {
    let out = run_main(
        "java.util.function.Predicate<Integer> even = n -> n % 2 == 0; System.out.println(even.test(4)); System.out.println(even.test(5));",
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn lambda_as_function_parses_strings_to_integers() {
    let out = run_main(
        "java.util.function.Function<String, Integer> parse = s -> Integer.parseInt(s); System.out.println(parse.apply(\"21\"));",
    );
    assert_eq!(out, vec!["21"]);
}

#[test]
fn lambda_as_consumer_prints_accepted_value() {
    let out = run_main(
        "java.util.function.Consumer<String> print = s -> System.out.println(s); print.accept(\"hi\");",
    );
    assert_eq!(out, vec!["hi"]);
}

#[test]
fn lambda_as_supplier_returns_constant() {
    let out = run_main(
        "java.util.function.Supplier<Integer> supply = () -> 42; System.out.println(supply.get());",
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn lambda_as_int_supplier_reads_captured_local() {
    let out = run_main(
        "int base = 10; java.util.function.IntSupplier supplier = () -> base + 1; System.out.println(supplier.getAsInt());",
    );
    assert_eq!(out, vec!["11"]);
}

#[test]
fn method_reference_to_static_parse_int() {
    let out = run_main(
        "java.util.function.Function<String, Integer> parser = Integer::parseInt; System.out.println(parser.apply(\"42\"));",
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn method_reference_to_string_length() {
    let out = run_main(
        "java.util.function.Function<String, Integer> len = String::length; System.out.println(len.apply(\"java\"));",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn method_reference_to_instance_to_upper_case() {
    let out = run_main(
        "java.util.function.Function<String, String> upper = String::toUpperCase; System.out.println(upper.apply(\"ok\"));",
    );
    assert_eq!(out, vec!["OK"]);
}

#[test]
fn lambda_passed_to_list_foreach() {
    let out = run_main("java.util.Arrays.asList(1, 2, 3).forEach(n -> System.out.println(n));");
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn predicate_negated_via_lambda_block() {
    let out = run_main(
        "java.util.function.Predicate<Integer> positive = n -> { return n > 0; }; System.out.println(positive.test(3)); System.out.println(positive.test(-1));",
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn function_composes_string_prefix() {
    let out = run_main(
        "java.util.function.Function<String, String> tag = s -> \"id:\" + s; System.out.println(tag.apply(\"7\"));",
    );
    assert_eq!(out, vec!["id:7"]);
}

#[test]
fn bi_function_adds_two_integers() {
    let out = run_main(
        "java.util.function.BiFunction<Integer, Integer, Integer> add = (a, b) -> a + b; System.out.println(add.apply(8, 5));",
    );
    assert_eq!(out, vec!["13"]);
}

#[test]
fn bi_consumer_prints_pair_of_values() {
    let out = run_main(
        "java.util.function.BiConsumer<Integer, Integer> pair = (a, b) -> System.out.println(a + b); pair.accept(2, 3);",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn int_predicate_detects_values_over_ten() {
    let out = run_main(
        "java.util.function.IntPredicate big = n -> n > 10; System.out.println(big.test(11)); System.out.println(big.test(9));",
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn int_function_squares_input() {
    let out = run_main(
        "java.util.function.IntFunction<Integer> square = n -> n * n; System.out.println(square.apply(5));",
    );
    assert_eq!(out, vec!["25"]);
}

#[test]
fn int_consumer_prints_doubled_value() {
    let out = run_main(
        "java.util.function.IntConsumer emit = n -> System.out.println(n * 2); emit.accept(6);",
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn lambda_used_in_stream_map() {
    let out = run_main(
        "java.util.List<Integer> doubled = java.util.Arrays.asList(1, 2, 3).stream().map(n -> n * 2).toList(); System.out.println(doubled.get(2));",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn lambda_used_in_stream_filter() {
    let out = run_main(
        "long count = java.util.Arrays.asList(1, 2, 3, 4).stream().filter(n -> n % 2 == 0).count(); System.out.println(count);",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn lambda_used_in_stream_reduce() {
    let out = run_main(
        "int sum = java.util.Arrays.asList(1, 2, 3).stream().reduce(0, (a, b) -> a + b); System.out.println(sum);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn supplier_factory_lambda_creates_incrementing_values() {
    let out = run_main(
        "int[] n = new int[]{0}; java.util.function.Supplier<Integer> next = () -> { n[0] = n[0] + 1; return n[0]; }; System.out.println(next.get()); System.out.println(next.get());",
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn consumer_accumulates_into_array_box() {
    let out = run_main(
        "int[] sum = new int[]{0}; java.util.function.Consumer<Integer> add = v -> { sum[0] = sum[0] + v; }; add.accept(4); add.accept(5); System.out.println(sum[0]);",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn function_identity_lambda_returns_argument() {
    let out = run_main(
        "java.util.function.Function<Integer, Integer> id = x -> x; System.out.println(id.apply(77));",
    );
    assert_eq!(out, vec!["77"]);
}

#[test]
fn predicate_and_composed_via_two_lambdas() {
    let out = run_main(
        "java.util.function.Predicate<Integer> gt2 = n -> n > 2; java.util.function.Predicate<Integer> lt5 = n -> n < 5; System.out.println(gt2.test(3)); System.out.println(lt5.test(3));",
    );
    assert_eq!(out, vec!["true", "true"]);
}

#[test]
fn method_reference_value_of_boxes_integer() {
    let out = run_main(
        "java.util.function.Function<Integer, String> fmt = String::valueOf; System.out.println(fmt.apply(15));",
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn lambda_with_explicit_return_in_block_body() {
    let out = run_main(
        "java.util.function.Function<Integer, Integer> inc = n -> { return n + 1; }; System.out.println(inc.apply(9));",
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn runnable_lambda_executes_print_side_effect() {
    let out = run_main("Runnable r = () -> System.out.println(\"run\"); r.run();");
    assert_eq!(out, vec!["run"]);
}

#[test]
fn callable_lambda_returns_computed_string() {
    let out = run_main(
        "java.util.concurrent.Callable<String> task = () -> \"done\"; System.out.println(task.call());",
    );
    assert_eq!(out, vec!["done"]);
}

#[test]
fn comparator_lambda_orders_integers_descending() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(3); list.add(2); list.sort((a, b) -> b - a); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["3", "1"]);
}

#[test]
fn unary_operator_on_strings_appends_suffix() {
    let out = run_main(
        "java.util.function.UnaryOperator<String> exclaim = s -> s + \"!\"; System.out.println(exclaim.apply(\"go\"));",
    );
    assert_eq!(out, vec!["go!"]);
}

#[test]
fn binary_operator_multiplies_two_integers() {
    let out = run_main(
        "java.util.function.BinaryOperator<Integer> mul = (a, b) -> a * b; System.out.println(mul.apply(6, 7));",
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn predicate_on_strings_checks_nonempty() {
    let out = run_main(
        "java.util.function.Predicate<String> nonEmpty = s -> s.length() > 0; System.out.println(nonEmpty.test(\"x\")); System.out.println(nonEmpty.test(\"\"));",
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn function_chained_through_two_lambda_assignments() {
    let out = run_main(
        "java.util.function.Function<Integer, Integer> step1 = n -> n + 1; java.util.function.Function<Integer, Integer> step2 = n -> step1.apply(n) * 2; System.out.println(step2.apply(3));",
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn method_reference_system_out_println_via_consumer() {
    let out = run_main(
        "java.util.function.Consumer<String> log = System.out::println; log.accept(\"logged\");",
    );
    assert_eq!(out, vec!["logged"]);
}

#[test]
fn lambda_captures_two_effectively_final_locals() {
    let out = run_main(
        "int a = 2; int b = 3; java.util.function.Supplier<Integer> sum = () -> a + b; System.out.println(sum.get());",
    );
    assert_eq!(out, vec!["5"]);
}
