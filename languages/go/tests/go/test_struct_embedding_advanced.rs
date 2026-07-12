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
    pointer_embedded_field_promotion_runtime => ("package main; import \"fmt\"; type inner struct { count int }; type outer struct { *inner }; func main() { value := outer{inner: &inner{count: 5}}; fmt.Println(value.count); }", vec!["5"]),
    pointer_embedded_nil_inner_runtime => ("package main; import \"fmt\"; type inner struct { count int }; type outer struct { *inner }; func main() { var value outer; fmt.Println(value.inner == nil); }", vec!["true"]),
    pointer_embedded_explicit_access_runtime => ("package main; import \"fmt\"; type inner struct { count int }; type outer struct { *inner }; func main() { value := outer{inner: &inner{count: 8}}; fmt.Println(value.inner.count); }", vec!["8"]),
    pointer_embedded_method_promotion_runtime => ("package main; import \"fmt\"; type inner struct { label string }; func (inner) name() string { return inner.label }; type outer struct { *inner }; func main() { value := outer{inner: &inner{label: \"vybe\"}}; fmt.Println(value.name()); }", vec!["vybe"]),
    pointer_embedded_promoted_bump_runtime => ("package main; import \"fmt\"; type inner struct { n int }; func (i *inner) bump() { i.n++ }; type outer struct { *inner }; func main() { value := outer{inner: &inner{n: 2}}; value.bump(); fmt.Println(value.n); }", vec!["3"]),
    multiple_anonymous_fields_promotion_runtime => ("package main; import \"fmt\"; type axis struct { x int }; type ord struct { y int }; type point struct { axis; ord }; func main() { value := point{axis: axis{x: 4}, ord: ord{y: 6}}; fmt.Println(value.x + value.y); }", vec!["10"]),
    outer_method_shadows_embedded_runtime => ("package main; import \"fmt\"; type inner struct{}; func (inner) label() string { return \"inner\" }; type outer struct { inner }; func (outer) label() string { return \"outer\" }; func main() { value := outer{}; fmt.Println(value.label()); }", vec!["outer"]),
    promoted_field_assignment_runtime => ("package main; import \"fmt\"; type inner struct { count int }; type outer struct { inner }; func main() { value := outer{inner: inner{count: 1}}; value.count = 9; fmt.Println(value.count); }", vec!["9"]),
    triple_nested_promoted_field_runtime => ("package main; import \"fmt\"; type leaf struct { value int }; type branch struct { leaf }; type trunk struct { branch }; func main() { value := trunk{branch: branch{leaf: leaf{value: 7}}}; fmt.Println(value.value); }", vec!["7"]),
    triple_nested_promoted_method_runtime => ("package main; import \"fmt\"; type leaf struct{}; func (leaf) tag() string { return \"deep\" }; type branch struct { leaf }; type trunk struct { branch }; func main() { value := trunk{}; fmt.Println(value.tag()); }", vec!["deep"]),
    nested_explicit_middle_field_runtime => ("package main; import \"fmt\"; type inner struct { count int }; type middle struct { inner }; type outer struct { middle }; func main() { value := outer{middle: middle{inner: inner{count: 11}}}; fmt.Println(value.middle.count); }", vec!["11"]),
    promoted_method_with_parameters_runtime => ("package main; import \"fmt\"; type inner struct { base int }; func (i inner) add(delta int) int { return i.base + delta }; type outer struct { inner }; func main() { value := outer{inner: inner{base: 3}}; fmt.Println(value.add(5)); }", vec!["8"]),
    promoted_method_value_runtime => ("package main; import \"fmt\"; type inner struct { n int }; func (i inner) total() int { return i.n }; type outer struct { inner }; func main() { value := outer{inner: inner{n: 6}}; fn := value.total; fmt.Println(fn()); }", vec!["6"]),
    dual_embedded_distinct_methods_runtime => ("package main; import \"fmt\"; type left struct{}; func (left) side() string { return \"L\" }; type right struct{}; func (right) edge() string { return \"R\" }; type pair struct { left; right }; func main() { value := pair{}; fmt.Println(value.side()); fmt.Println(value.edge()); }", vec!["L", "R"]),
    promoted_field_survives_struct_copy_runtime => ("package main; import \"fmt\"; type inner struct { count int }; type outer struct { inner }; func main() { original := outer{inner: inner{count: 2}}; copy := original; copy.count = 5; fmt.Println(original.count); fmt.Println(copy.count); }", vec!["2", "5"]),
    four_level_field_promotion_runtime => ("package main; import \"fmt\"; type d struct { n int }; type c struct { d }; type b struct { c }; type a struct { b }; func main() { value := a{b: b{c: c{d: d{n: 13}}}}; fmt.Println(value.n); }", vec!["13"]),
    value_embed_pointer_receiver_promotion_runtime => ("package main; import \"fmt\"; type inner struct { n int }; func (i *inner) double() { i.n = i.n * 2 }; type outer struct { inner }; func main() { value := outer{inner: inner{n: 4}}; value.double(); fmt.Println(value.n); }", vec!["8"]),
    nested_pointer_middle_embedding_runtime => ("package main; import \"fmt\"; type inner struct { count int }; type middle struct { *inner }; type outer struct { middle }; func main() { value := outer{middle: middle{inner: &inner{count: 15}}}; fmt.Println(value.count); }", vec!["15"]),
    promoted_string_field_concat_runtime => ("package main; import \"fmt\"; type inner struct { prefix string; suffix string }; type outer struct { inner }; func main() { value := outer{inner: inner{prefix: \"go\", suffix: \"lang\"}}; fmt.Println(value.prefix + value.suffix); }", vec!["golang"]),
    embedded_bool_promoted_in_condition_runtime => ("package main; import \"fmt\"; type inner struct { ready bool }; type outer struct { inner }; func main() { value := outer{inner: inner{ready: true}}; if value.ready { fmt.Println(1) } else { fmt.Println(0) } }", vec!["1"]),
}

compile_cases! {
    pointer_embedded_literal_compile => "package main; type inner struct { count int }; type outer struct { *inner }; func main() { _ = outer{inner: &inner{count: 1}} }",
    multiple_anonymous_fields_compile => "package main; type axis struct { x int }; type ord struct { y int }; type point struct { axis; ord }; func main() { var value point; _ = value.x; _ = value.y }",
    outer_method_shadows_embedded_compile => "package main; type inner struct{}; func (inner) label() string { return \"inner\" }; type outer struct { inner }; func (outer) label() string { return \"outer\" }; func main() { _ = outer{}.label() }",
    triple_nested_embedding_compile => "package main; type leaf struct { value int }; type branch struct { leaf }; type trunk struct { branch }; func main() { var value trunk; _ = value.value }",
    nested_pointer_middle_embedding_compile => "package main; type inner struct { count int }; type middle struct { *inner }; type outer struct { middle }; func main() { var value outer; _ = value.count }",
    promoted_pointer_receiver_method_compile => "package main; type inner struct { n int }; func (i *inner) bump() { i.n++ }; type outer struct { inner }; func main() { var value outer; value.bump() }",
    address_of_promoted_field_compile => "package main; type inner struct { count int }; type outer struct { inner }; func main() { var value outer; ptr := &value.count; _ = ptr }",
    dual_embedded_distinct_methods_compile => "package main; type left struct{}; func (left) side() string { return \"L\" }; type right struct{}; func (right) edge() string { return \"R\" }; type pair struct { left; right }; func main() { _ = pair{}.side(); _ = pair{}.edge() }",
    deep_method_promotion_compile => "package main; type leaf struct{}; func (leaf) tag() string { return \"deep\" }; type branch struct { leaf }; type trunk struct { branch }; func main() { _ = trunk{}.tag() }",
    pointer_embed_field_selector_compile => "package main; type inner struct { count int }; type outer struct { *inner }; func main() { value := outer{inner: &inner{count: 2}}; _ = value.count }",
}
