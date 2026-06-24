use crate::helpers::run_main;

#[test]
fn lambda_expression_as_functional_interface() {
    let out = run_main(
        "java.util.function.IntUnaryOperator doubler = x -> x * 2; System.out.println(doubler.applyAsInt(6));",
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn method_reference_to_static_method() {
    let out = run_main(
        "java.util.function.Function<String, Integer> parser = Integer::parseInt; System.out.println(parser.apply(\"42\"));",
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn lambda_captures_effectively_final_local() {
    let out = run_main(
        "int base = 10; java.util.function.IntSupplier supplier = () -> base + 1; System.out.println(supplier.getAsInt());",
    );
    assert_eq!(out, vec!["11"]);
}

#[test]
fn stream_map_collects_transformed_values() {
    let out = run_main(
        "java.util.List<Integer> nums = java.util.Arrays.asList(1, 2, 3); java.util.List<Integer> doubled = nums.stream().map(n -> n * 2).toList(); System.out.println(doubled.get(2));",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn stream_filter_keeps_matching_elements() {
    let out = run_main(
        "long count = java.util.Arrays.asList(1, 2, 3, 4).stream().filter(n -> n % 2 == 0).count(); System.out.println(count);",
    );
    assert_eq!(out, vec!["2"]);
}
