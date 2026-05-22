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
    empty_interface_holds_int_runtime => ("package main; import \"fmt\"; func main() { var value interface{} = 7; fmt.Println(value); }", vec!["7"]),
    empty_interface_holds_string_runtime => ("package main; import \"fmt\"; func main() { var value interface{} = \"go\"; fmt.Println(value); }", vec!["go"]),
    interface_method_dispatch_runtime => ("package main; import \"fmt\"; type speaker interface { speak() string }; type dog struct{}; func (dog) speak() string { return \"woof\" }; func main() { var value speaker = dog{}; fmt.Println(value.speak()); }", vec!["woof"]),
    interface_reassignment_runtime => ("package main; import \"fmt\"; func main() { var value interface{} = 1; fmt.Println(value); value = \"two\"; fmt.Println(value); }", vec!["1", "two"]),
    interface_slice_element_print_runtime => ("package main; import \"fmt\"; func main() { values := []interface{}{1, \"go\"}; fmt.Println(values[1]); }", vec!["go"]),
    interface_map_value_print_runtime => ("package main; import \"fmt\"; func main() { values := map[string]interface{}{\"n\": 4}; fmt.Println(values[\"n\"]); }", vec!["4"]),
    interface_nil_compare_after_assignment_runtime => ("package main; import \"fmt\"; func main() { var value interface{}; fmt.Println(value == nil); value = 3; fmt.Println(value == nil); }", vec!["true", "false"]),
    interface_with_pointer_receiver_runtime => ("package main; import \"fmt\"; type sized interface { size() int }; type box struct { n int }; func (b *box) size() int { return b.n }; func main() { var value sized = &box{n: 6}; fmt.Println(value.size()); }", vec!["6"]),
    interface_return_value_runtime => ("package main; import \"fmt\"; type speaker interface { speak() string }; type cat struct{}; func (cat) speak() string { return \"meow\" }; func build() speaker { return cat{} }; func main() { fmt.Println(build().speak()); }", vec!["meow"]),
    interface_pass_as_parameter_runtime => ("package main; import \"fmt\"; type speaker interface { speak() string }; type bird struct{}; func (bird) speak() string { return \"chirp\" }; func say(value speaker) string { return value.speak() }; func main() { fmt.Println(say(bird{})); }", vec!["chirp"]),
    interface_field_in_struct_runtime => ("package main; import \"fmt\"; type speaker interface { speak() string }; type dog struct{}; func (dog) speak() string { return \"woof\" }; type holder struct { value speaker }; func main() { h := holder{value: dog{}}; fmt.Println(h.value.speak()); }", vec!["woof"]),
    interface_value_in_slice_runtime => ("package main; import \"fmt\"; type speaker interface { speak() string }; type dog struct{}; func (dog) speak() string { return \"woof\" }; func main() { values := []speaker{dog{}}; fmt.Println(values[0].speak()); }", vec!["woof"]),
    interface_value_in_map_runtime => ("package main; import \"fmt\"; type speaker interface { speak() string }; type dog struct{}; func (dog) speak() string { return \"woof\" }; func main() { values := map[string]speaker{\"pet\": dog{}}; fmt.Println(values[\"pet\"].speak()); }", vec!["woof"]),
    interface_multiple_implementers_runtime => ("package main; import \"fmt\"; type speaker interface { speak() string }; type dog struct{}; type cat struct{}; func (dog) speak() string { return \"woof\" }; func (cat) speak() string { return \"meow\" }; func main() { values := []speaker{dog{}, cat{}}; fmt.Println(values[0].speak()); fmt.Println(values[1].speak()); }", vec!["woof", "meow"]),
    empty_interface_compare_nil_runtime => ("package main; import \"fmt\"; func main() { var value interface{}; fmt.Println(value == nil); }", vec!["true"]),
    interface_method_returns_string_runtime => ("package main; import \"fmt\"; type namer interface { name() string }; type widget struct{}; func (widget) name() string { return \"vybe\" }; func main() { var value namer = widget{}; fmt.Println(value.name()); }", vec!["vybe"]),
    interface_zero_value_field_runtime => ("package main; import \"fmt\"; type holder struct { value interface{} }; func main() { var h holder; fmt.Println(h.value == nil); }", vec!["true"]),
    interface_value_roundtrip_runtime => ("package main; import \"fmt\"; func wrap(v interface{}) interface{} { return v }; func main() { fmt.Println(wrap(9)); }", vec!["9"]),
    interface_method_with_struct_receiver_runtime => ("package main; import \"fmt\"; type sized interface { size() int }; type box struct { n int }; func (b box) size() int { return b.n }; func main() { var value sized = box{n: 8}; fmt.Println(value.size()); }", vec!["8"]),
    empty_interface_fmt_print_runtime => ("package main; import \"fmt\"; func main() { values := []interface{}{3, \"go\"}; fmt.Println(values[0]); fmt.Println(values[1]); }", vec!["3", "go"]),
    interface_assignment_between_variables_runtime => ("package main; import \"fmt\"; func main() { var left interface{} = \"go\"; right := left; fmt.Println(right); }", vec!["go"]),
    empty_interface_bool_runtime => ("package main; import \"fmt\"; func main() { var value interface{} = true; fmt.Println(value); }", vec!["true"]),
    empty_interface_array_element_runtime => ("package main; import \"fmt\"; func main() { values := [2]interface{}{\"a\", 2}; fmt.Println(values[0]); fmt.Println(values[1]); }", vec!["a", "2"]),
    interface_from_function_return_runtime => ("package main; import \"fmt\"; func build() interface{} { return \"built\" }; func main() { fmt.Println(build()); }", vec!["built"]),
    interface_in_struct_literal_runtime => ("package main; import \"fmt\"; type holder struct { value interface{} }; func main() { value := holder{value: 11}; fmt.Println(value.value); }", vec!["11"]),
}

compile_cases! {
    interface_type_assertion_compile => "package main; func main() { var value interface{} = 1; _ = value.(int) }",
    interface_type_switch_compile => "package main; func main() { var value interface{} = 1; switch value.(type) { case int: } }",
    interface_embedding_compile => "package main; type reader interface { read() int }; type closer interface { close() }; type resource interface { reader; closer }; func main() {}",
    empty_interface_in_map_compile => "package main; func main() { values := map[string]interface{}{\"x\": 1}; _ = values }",
    interface_in_slice_compile => "package main; type reader interface { read() int }; func main() { var values []reader; _ = values }",
    interface_in_struct_compile => "package main; type reader interface { read() int }; type holder struct { value reader }; func main() { _ = holder{} }",
    interface_method_set_pointer_compile => "package main; type reader interface { read() int }; type box struct{}; func (b *box) read() int { return 1 }; func main() { var value reader = &box{}; _ = value }",
    interface_method_set_value_compile => "package main; type reader interface { read() int }; type box struct{}; func (b box) read() int { return 1 }; func main() { var value reader = box{}; _ = value }",
    interface_returning_interface_compile => "package main; type reader interface { read() int }; type box struct{}; func (b box) read() int { return 1 }; func build() reader { return box{} }; func main() { _ = build() }",
    interface_parameter_compile => "package main; type reader interface { read() int }; func use(value reader) int { return value.read() }; func main() { _ = use }",
    interface_assignment_compile => "package main; func main() { var left interface{} = 1; var right interface{} = left; _ = right }",
    named_interface_alias_compile => "package main; type reader interface { read() int }; type alias = reader; func main() { var value alias; _ = value }",
    interface_with_multiple_methods_compile => "package main; type shape interface { area() int; perimeter() int }; func main() {}",
    interface_with_blank_identifier_assignment_compile => "package main; func main() { var value interface{} = 1; _ = value }",
    interface_slice_of_empty_interface_compile => "package main; func main() { values := []interface{}{1, \"go\"}; _ = values }",
    interface_map_key_compile => "package main; func main() { values := map[interface{}]string{1: \"one\"}; _ = values }",
    interface_method_returning_interface_compile => "package main; type builder interface { build() interface{} }; func main() {}",
    interface_with_pointer_argument_compile => "package main; type loader interface { load(*int) }; func main() {}",
    interface_in_variadic_compile => "package main; func pack(values ...interface{}) []interface{} { return values }; func main() { _ = pack(1, \"two\") }",
    type_assertion_two_result_compile => "package main; func main() { var value interface{} = 1; number, ok := value.(int); _, _ = number, ok }",
    type_switch_with_short_decl_compile => "package main; func main() { switch value := interface{}(1); value.(type) { case int: } }",
    interface_composite_literal_field_compile => "package main; type holder struct { value interface{} }; func main() { _ = holder{value: 1} }",
    interface_zero_value_compile => "package main; func main() { var value interface{}; _ = value }",
    interface_compare_nil_compile => "package main; func main() { var value interface{}; _ = (value == nil) }",
    interface_array_compile => "package main; func main() { values := [2]interface{}{1, \"go\"}; _ = values }",
}