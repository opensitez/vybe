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
    value_receiver_reads_field_runtime => ("package main; import \"fmt\"; type counter struct { n int }; func (c counter) total() int { return c.n }; func main() { value := counter{n: 5}; fmt.Println(value.total()); }", vec!["5"]),
    pointer_receiver_updates_field_runtime => ("package main; import \"fmt\"; type counter struct { n int }; func (c *counter) add(v int) { c.n += v }; func main() { value := counter{n: 2}; value.add(4); fmt.Println(value.n); }", vec!["6"]),
    method_value_invocation_runtime => ("package main; import \"fmt\"; type counter struct { n int }; func (c counter) total() int { return c.n }; func main() { value := counter{n: 7}; fn := value.total; fmt.Println(fn()); }", vec!["7"]),
    method_returns_string_runtime => ("package main; import \"fmt\"; type label struct { text string }; func (l label) value() string { return l.text }; func main() { fmt.Println(label{text: \"vybe\"}.value()); }", vec!["vybe"]),
    method_with_multiple_params_runtime => ("package main; import \"fmt\"; type counter struct { n int }; func (c counter) add(a int, b int) int { return c.n + a + b }; func main() { value := counter{n: 1}; fmt.Println(value.add(2, 3)); }", vec!["6"]),
    method_on_named_type_runtime => ("package main; import \"fmt\"; type score int; func (s score) next() int { return int(s) + 1 }; func main() { var value score = 8; fmt.Println(value.next()); }", vec!["9"]),
    pointer_method_on_new_runtime => ("package main; import \"fmt\"; type counter struct { n int }; func (c *counter) bump() { c.n++ }; func main() { value := new(counter); value.bump(); fmt.Println(value.n); }", vec!["1"]),
    value_receiver_copy_does_not_mutate_runtime => ("package main; import \"fmt\"; type counter struct { n int }; func (c counter) bump() { c.n++ }; func main() { value := counter{n: 3}; value.bump(); fmt.Println(value.n); }", vec!["3"]),
    method_result_used_in_expression_runtime => ("package main; import \"fmt\"; type counter struct { n int }; func (c counter) total() int { return c.n }; func main() { value := counter{n: 4}; fmt.Println(value.total() * 2); }", vec!["8"]),
    method_on_struct_parameter_runtime => ("package main; import \"fmt\"; type counter struct { n int }; func (c counter) total() int { return c.n }; func show(value counter) int { return value.total() }; func main() { fmt.Println(show(counter{n: 10})); }", vec!["10"]),
    method_on_embedded_field_explicit_runtime => ("package main; import \"fmt\"; type inner struct{}; func (inner) label() string { return \"ok\" }; type outer struct { inner }; func main() { value := outer{}; fmt.Println(value.inner.label()); }", vec!["ok"]),
    pointer_receiver_called_twice_runtime => ("package main; import \"fmt\"; type counter struct { n int }; func (c *counter) add(v int) { c.n += v }; func main() { value := counter{}; value.add(2); value.add(5); fmt.Println(value.n); }", vec!["7"]),
    method_with_array_field_runtime => ("package main; import \"fmt\"; type bag struct { values [2]int }; func (b bag) second() int { return b.values[1] }; func main() { value := bag{values: [2]int{3, 9}}; fmt.Println(value.second()); }", vec!["9"]),
    method_with_slice_field_runtime => ("package main; import \"fmt\"; type bag struct { values []int }; func (b bag) count() int { return len(b.values) }; func main() { value := bag{values: []int{1, 2, 3}}; fmt.Println(value.count()); }", vec!["3"]),
    method_with_map_field_runtime => ("package main; import \"fmt\"; type bag struct { values map[string]int }; func (b bag) get(key string) int { return b.values[key] }; func main() { value := bag{values: map[string]int{\"a\": 6}}; fmt.Println(value.get(\"a\")); }", vec!["6"]),
    zero_value_method_runtime => ("package main; import \"fmt\"; type counter struct { n int }; func (c counter) total() int { return c.n }; func main() { var value counter; fmt.Println(value.total()); }", vec!["0"]),
    method_in_if_condition_runtime => ("package main; import \"fmt\"; type gate struct { open bool }; func (g gate) ready() bool { return g.open }; func main() { value := gate{open: true}; if value.ready() { fmt.Println(1) } else { fmt.Println(0) } }", vec!["1"]),
    method_return_bool_runtime => ("package main; import \"fmt\"; type gate struct { open bool }; func (g gate) ready() bool { return g.open }; func main() { fmt.Println(gate{open: false}.ready()); }", vec!["false"]),
    method_on_alias_like_named_type_runtime => ("package main; import \"fmt\"; type text string; func (t text) label() string { return string(t) + \"!\" }; func main() { var value text = \"go\"; fmt.Println(value.label()); }", vec!["go!"]),
    method_returns_struct_field_runtime => ("package main; import \"fmt\"; type point struct { x int }; func (p point) value() int { return p.x }; func main() { fmt.Println(point{x: 12}.value()); }", vec!["12"]),
    pointer_receiver_through_alias_variable_runtime => ("package main; import \"fmt\"; type counter struct { n int }; func (c *counter) add(v int) { c.n += v }; func main() { value := counter{n: 1}; alias := &value; alias.add(8); fmt.Println(value.n); }", vec!["9"]),
    method_value_from_pointer_receiver_runtime => ("package main; import \"fmt\"; type counter struct { n int }; func (c *counter) bump() { c.n++ }; func main() { value := &counter{n: 4}; fn := value.bump; fn(); fmt.Println(value.n); }", vec!["5"]),
    method_on_literal_runtime => ("package main; import \"fmt\"; type counter struct { n int }; func (c counter) total() int { return c.n }; func main() { fmt.Println(counter{n: 14}.total()); }", vec!["14"]),
    method_on_returned_struct_runtime => ("package main; import \"fmt\"; type counter struct { n int }; func (c counter) total() int { return c.n }; func build() counter { return counter{n: 15} }; func main() { fmt.Println(build().total()); }", vec!["15"]),
    method_chosen_from_variable_runtime => ("package main; import \"fmt\"; type counter struct { n int }; func (c counter) total() int { return c.n }; func main() { value := counter{n: 16}; other := value; fmt.Println(other.total()); }", vec!["16"]),
}

compile_cases! {
    method_expression_value_receiver_compile => "package main; type counter struct { n int }; func (c counter) total() int { return c.n }; func main() { _ = counter.total }",
    method_expression_pointer_receiver_compile => "package main; type counter struct { n int }; func (c *counter) bump() { c.n++ }; func main() { _ = (*counter).bump }",
    pointer_receiver_interface_assign_compile => "package main; type adder interface { add(int) }; type counter struct { n int }; func (c *counter) add(v int) { c.n += v }; func main() { var value adder = &counter{}; _ = value }",
    value_receiver_interface_assign_compile => "package main; type totaler interface { total() int }; type counter struct { n int }; func (c counter) total() int { return c.n }; func main() { var value totaler = counter{}; _ = value }",
    method_on_empty_struct_compile => "package main; type marker struct{}; func (marker) ok() bool { return true }; func main() { _ = marker{}.ok() }",
    embedded_method_compile => "package main; type inner struct{}; func (inner) label() string { return \"ok\" }; type outer struct { inner }; func main() { var value outer; _ = value.inner.label() }",
    pointer_method_on_struct_literal_compile => "package main; type counter struct { n int }; func (c *counter) bump() { c.n++ }; func main() { value := &counter{}; value.bump() }",
    method_with_named_return_compile => "package main; type counter struct { n int }; func (c counter) total() (result int) { result = c.n; return }; func main() { _ = counter{}.total() }",
    method_returning_pointer_compile => "package main; type node struct { next *node }; func (n node) clone() *node { return &node{} }; func main() { _ = node{}.clone() }",
    method_value_assignment_compile => "package main; type counter struct { n int }; func (c counter) total() int { return c.n }; func main() { value := counter{}; fn := value.total; _ = fn }",
    pointer_method_value_assignment_compile => "package main; type counter struct { n int }; func (c *counter) bump() { c.n++ }; func main() { value := &counter{}; fn := value.bump; _ = fn }",
    method_on_named_type_compile => "package main; type score int; func (s score) next() int { return int(s) + 1 }; func main() { var value score; _ = value.next() }",
    method_with_slice_field_compile => "package main; type bag struct { values []int }; func (b bag) count() int { return len(b.values) }; func main() { _ = bag{}.count() }",
    method_with_map_field_compile => "package main; type bag struct { values map[string]int }; func (b bag) size() int { return len(b.values) }; func main() { _ = bag{}.size() }",
    method_with_array_field_compile => "package main; type bag struct { values [2]int }; func (b bag) first() int { return b.values[0] }; func main() { _ = bag{}.first() }",
    method_used_as_callback_compile => "package main; type counter struct { n int }; func (c counter) total() int { return c.n }; func main() { value := counter{}; fn := value.total; _ = fn() }",
    method_in_struct_field_compile => "package main; type counter struct { n int }; func (c counter) total() int { return c.n }; type holder struct { fn func() int }; func main() { value := counter{}; _ = holder{fn: value.total} }",
    method_call_on_function_return_compile => "package main; type counter struct { n int }; func (c counter) total() int { return c.n }; func build() counter { return counter{} }; func main() { _ = build().total() }",
    pointer_receiver_chain_compile => "package main; type counter struct { n int }; func (c *counter) bump() { c.n++ }; func main() { value := &counter{}; value.bump(); value.bump() }",
    method_with_multiple_params_compile => "package main; type counter struct { n int }; func (c counter) add(a int, b int) int { return c.n + a + b }; func main() { _ = counter{}.add(1, 2) }",
    method_with_bool_return_compile => "package main; type gate struct { open bool }; func (g gate) ready() bool { return g.open }; func main() { _ = gate{}.ready() }",
    method_on_alias_like_named_type_compile => "package main; type text string; func (t text) label() string { return string(t) }; func main() { var value text; _ = value.label() }",
    method_receiver_shadow_compile => "package main; type counter struct { n int }; func (c counter) total() int { c := counter{n: c.n + 1}; return c.n }; func main() { _ = counter{}.total() }",
    pointer_receiver_on_new_compile => "package main; type counter struct { n int }; func (c *counter) bump() { c.n++ }; func main() { value := new(counter); value.bump() }",
    method_returning_struct_compile => "package main; type point struct { x int }; type builder struct{}; func (builder) build() point { return point{x: 1} }; func main() { _ = builder{}.build() }",
}