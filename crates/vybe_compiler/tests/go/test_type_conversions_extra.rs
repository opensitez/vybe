use crate::helpers::*;

macro_rules! go_compile_test {
    ($name:ident, $src:expr) => {
        #[test]
        fn $name() {
            compile_ok($src);
        }
    };
}

macro_rules! go_run_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            let out = run_prints($src);
            assert_eq!(out, $expected);
        }
    };
}

macro_rules! run_cases {
    ($( $name:ident => ($src:expr, $expected:expr), )*) => {
        $( go_run_test!($name, $src, $expected); )*
    };
}

macro_rules! compile_cases {
    ($( $name:ident => $src:expr, )*) => {
        $( go_compile_test!($name, $src); )*
    };
}

run_cases! {
    int_to_float_runtime => ("package main; import \"fmt\"; func main() { fmt.Println(float64(3)); }", vec!["3"]),
    float_to_int_trunc_runtime => ("package main; import \"fmt\"; func main() { fmt.Println(int(3.9)); }", vec!["3"]),
    named_int_to_int_runtime => ("package main; import \"fmt\"; type score int; func main() { var value score = 7; fmt.Println(int(value)); }", vec!["7"]),
    int_to_named_type_runtime => ("package main; import \"fmt\"; type score int; func main() { value := score(8); fmt.Println(value); }", vec!["8"]),
    alias_type_conversion_runtime => ("package main; import \"fmt\"; type count = int; func main() { value := count(9); fmt.Println(value); }", vec!["9"]),
    byte_to_int_runtime => ("package main; import \"fmt\"; func main() { var value byte = 10; fmt.Println(int(value)); }", vec!["10"]),
    rune_to_int_runtime => ("package main; import \"fmt\"; func main() { var value rune = 'A'; fmt.Println(int(value)); }", vec!["65"]),
    int_to_rune_runtime => ("package main; import \"fmt\"; func main() { fmt.Println(rune(66)); }", vec!["66"]),
    conversion_in_binary_runtime => ("package main; import \"fmt\"; func main() { fmt.Println(int(2.5) + 4); }", vec!["6"]),
    conversion_in_return_runtime => ("package main; import \"fmt\"; func cast(v int) float64 { return float64(v) }; func main() { fmt.Println(cast(5)); }", vec!["5"]),
    slice_element_conversion_runtime => ("package main; import \"fmt\"; func main() { values := []int{4, 6}; fmt.Println(float64(values[1])); }", vec!["6"]),
    struct_field_conversion_runtime => ("package main; import \"fmt\"; type holder struct { count int }; func main() { value := holder{count: 12}; fmt.Println(float64(value.count)); }", vec!["12"]),
    map_value_conversion_runtime => ("package main; import \"fmt\"; func main() { values := map[string]int{\"a\": 13}; fmt.Println(float64(values[\"a\"])); }", vec!["13"]),
    array_value_conversion_runtime => ("package main; import \"fmt\"; func main() { values := [2]int{1, 14}; fmt.Println(float64(values[1])); }", vec!["14"]),
    float32_to_float64_runtime => ("package main; import \"fmt\"; func main() { var value float32 = 15; fmt.Println(float64(value)); }", vec!["15"]),
    int_literal_named_type_runtime => ("package main; import \"fmt\"; type level int; func main() { value := level(16); fmt.Println(value); }", vec!["16"]),
    method_on_named_converted_value_runtime => ("package main; import \"fmt\"; type level int; func (l level) next() int { return int(l) + 1 }; func main() { fmt.Println(level(17).next()); }", vec!["18"]),
    conversion_in_short_decl_runtime => ("package main; import \"fmt\"; func main() { value := int(18.2); fmt.Println(value); }", vec!["18"]),
    conversion_between_named_types_same_underlying_runtime => ("package main; import \"fmt\"; type first int; type second int; func main() { var value first = 19; fmt.Println(second(value)); }", vec!["19"]),
    unary_negation_after_conversion_runtime => ("package main; import \"fmt\"; func main() { fmt.Println(-int(20.9)); }", vec!["-20"]),
    conversion_in_if_compare_runtime => ("package main; import \"fmt\"; func main() { if int(21.1) == 21 { fmt.Println(1) } else { fmt.Println(0) } }", vec!["1"]),
    conversion_of_zero_value_named_type_runtime => ("package main; import \"fmt\"; type score int; func main() { var value score; fmt.Println(int(value)); }", vec!["0"]),
    nested_conversion_runtime => ("package main; import \"fmt\"; func main() { fmt.Println(int(float64(22))); }", vec!["22"]),
    conversion_inside_function_call_runtime => ("package main; import \"fmt\"; func show(v float64) { fmt.Println(v) }; func main() { show(float64(23)); }", vec!["23"]),
    conversion_from_named_type_in_expression_runtime => ("package main; import \"fmt\"; type score int; func main() { var value score = 24; fmt.Println(int(value) + 1); }", vec!["25"]),
}

compile_cases! {
    named_type_conversion_compile => "package main; type score int; func main() { _ = score(1) }",
    alias_type_conversion_compile => "package main; type count = int; func main() { _ = count(2) }",
    float_to_int_compile => "package main; func main() { _ = int(3.5) }",
    int_to_float_compile => "package main; func main() { _ = float64(4) }",
    rune_conversion_compile => "package main; func main() { _ = rune(65) }",
    byte_conversion_compile => "package main; func main() { _ = byte(6) }",
    nested_conversion_compile => "package main; func main() { _ = int(float64(7)) }",
    conversion_in_return_compile => "package main; func cast(v int) float64 { return float64(v) }; func main() { _ = cast }",
    conversion_in_assignment_compile => "package main; func main() { var value int; value = int(8.9); _ = value }",
    conversion_in_var_decl_compile => "package main; var value = int(9.4); func main() { _ = value }",
    named_to_named_conversion_compile => "package main; type first int; type second int; func main() { var value first = 10; _ = second(value) }",
    conversion_in_slice_literal_compile => "package main; func main() { _ = []float64{float64(1), float64(2)} }",
    conversion_in_array_literal_compile => "package main; func main() { _ = [2]int{int(1.2), int(2.3)} }",
    conversion_in_map_literal_compile => "package main; func main() { _ = map[string]float64{\"a\": float64(3)} }",
    conversion_in_struct_literal_compile => "package main; type holder struct { value float64 }; func main() { _ = holder{value: float64(4)} }",
    method_on_named_type_compile => "package main; type score int; func (s score) next() int { return int(s) + 1 }; func main() { _ = score(5).next() }",
    conversion_of_index_expression_compile => "package main; func main() { values := []int{6}; _ = float64(values[0]) }",
    conversion_of_map_lookup_compile => "package main; func main() { values := map[string]int{\"a\": 7}; _ = float64(values[\"a\"]) }",
    conversion_of_struct_field_compile => "package main; type holder struct { value int }; func main() { h := holder{value: 8}; _ = float64(h.value) }",
    conversion_in_if_init_compile => "package main; func main() { if value := int(9.8); value > 0 { _ = value } }",
    conversion_in_switch_init_compile => "package main; func main() { switch value := int(10.3); value { case 10: } }",
    conversion_in_for_clause_compile => "package main; func main() { for i := int(0); i < 1; i++ { _ = i } }",
    conversion_to_named_alias_compile => "package main; type level int; func main() { _ = level(int(11)) }",
    conversion_from_named_alias_compile => "package main; type level int; func main() { var value level = 12; _ = int(value) }",
    conversion_with_unary_minus_compile => "package main; func main() { _ = -int(13.7) }",
}