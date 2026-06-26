//! Type aliases (`type T = U`) vs defined types (`type T U`): identity,
//! struct fields, method sets, and underlying-type conversions.

use crate::helpers::*;

go_run_cases! {
    // --- alias identical to underlying (no conversion) ---
    alias_assign_from_untyped_literal => (
        "package main; import \"fmt\"; type Count = int; func main() { var value Count = 7; fmt.Println(value) }",
        vec!["7"]
    ),
    alias_assign_to_builtin_without_cast => (
        "package main; import \"fmt\"; type Count = int; func main() { var count Count = 8; var plain int = count; fmt.Println(plain) }",
        vec!["8"]
    ),
    alias_from_builtin_without_cast => (
        "package main; import \"fmt\"; type Count = int; func main() { plain := 9; var count Count = plain; fmt.Println(count) }",
        vec!["9"]
    ),
    alias_compare_equal_to_underlying => (
        "package main; import \"fmt\"; type Count = int; func main() { count := Count(10); fmt.Println(count == 10) }",
        vec!["true"]
    ),

    // --- defined type distinct from underlying (explicit conversion) ---
    defined_assign_from_untyped_literal => (
        "package main; import \"fmt\"; type Score int; func main() { var value Score = 12; fmt.Println(value) }",
        vec!["12"]
    ),
    defined_to_underlying_explicit_cast => (
        "package main; import \"fmt\"; type Score int; func main() { value := Score(13); fmt.Println(int(value)) }",
        vec!["13"]
    ),
    underlying_to_defined_explicit_cast => (
        "package main; import \"fmt\"; type Score int; func main() { fmt.Println(Score(14)) }",
        vec!["14"]
    ),
    defined_types_same_underlying_mutual_cast => (
        "package main; import \"fmt\"; type First int; type Second int; func main() { value := Second(First(15)); fmt.Println(value) }",
        vec!["15"]
    ),
    defined_cast_in_arithmetic_expression => (
        "package main; import \"fmt\"; type Score int; func main() { var value Score = 16; fmt.Println(int(value) + 1) }",
        vec!["17"]
    ),

    // --- alias and defined types in struct fields ---
    struct_field_with_int_alias => (
        "package main; import \"fmt\"; type Count = int; type row struct { total Count }; func main() { value := row{total: 18}; fmt.Println(value.total) }",
        vec!["18"]
    ),
    struct_field_with_slice_alias => (
        "package main; import \"fmt\"; type IDs = []int; type batch struct { items IDs }; func main() { value := batch{items: IDs{3, 4}}; fmt.Println(len(value.items)); fmt.Println(value.items[1]) }",
        vec!["2", "4"]
    ),
    struct_field_with_defined_string_type => (
        "package main; import \"fmt\"; type Label string; type item struct { name Label }; func main() { value := item{name: \"vybe\"}; fmt.Println(value.name) }",
        vec!["vybe"]
    ),
    struct_literal_defined_field_from_conversion => (
        "package main; import \"fmt\"; type Meters int; type segment struct { length Meters }; func main() { value := segment{length: Meters(19)}; fmt.Println(int(value.length)) }",
        vec!["19"]
    ),

    // --- methods on defined types; alias inherits defined method set ---
    defined_int_method_on_value_receiver => (
        "package main; import \"fmt\"; type Meters int; func (m Meters) Double() int { return int(m) * 2 }; func main() { fmt.Println(Meters(5).Double()) }",
        vec!["10"]
    ),
    defined_string_method_on_value_receiver => (
        "package main; import \"fmt\"; type Tag string; func (t Tag) Len() int { return len(string(t)) }; func main() { fmt.Println(Tag(\"go\").Len()) }",
        vec!["2"]
    ),
    defined_type_value_receiver_returns_new_value => (
        "package main; import \"fmt\"; type Offset int; func (o Offset) next() Offset { return o + 1 }; func main() { fmt.Println(Offset(2).next()) }",
        vec!["3"]
    ),
    alias_same_underlying_as_defined_without_conversion => (
        "package main; import \"fmt\"; type Units int; type Reading = Units; func main() { var base Units = 4; var view Reading = base; fmt.Println(int(view)) }",
        vec!["4"]
    ),

    // --- underlying-type conversions in collections and zero values ---
    slice_of_defined_type_element_cast => (
        "package main; import \"fmt\"; type Score int; func main() { values := []Score{Score(1), Score(2)}; fmt.Println(int(values[1])) }",
        vec!["2"]
    ),
    map_with_defined_type_value_cast => (
        "package main; import \"fmt\"; type Level int; func main() { values := map[string]Level{\"a\": Level(20)}; fmt.Println(int(values[\"a\"])) }",
        vec!["20"]
    ),
    zero_value_defined_type_prints_zero => (
        "package main; import \"fmt\"; type Score int; func main() { var value Score; fmt.Println(int(value)) }",
        vec!["0"]
    ),
}

go_compile_cases! {
    alias_map_type_in_struct_field => "package main; type Dict = map[string]int; type holder struct { data Dict }; func main() { _ = holder{data: Dict{\"k\": 1}} }",
    alias_pointer_type_in_struct_field => "package main; type IntPtr = *int; type holder struct { ptr IntPtr }; func main() { n := 1; _ = holder{ptr: &n} }",
    method_on_defined_struct_type => "package main; type Row struct { n int }; func (r Row) total() int { return r.n }; func main() { _ = Row{n: 1}.total() }",
    alias_inherits_methods_from_defined_target => "package main; type Units int; func (u Units) sign() int { if u < 0 { return -1 }; if u > 0 { return 1 }; return 0 }; type Reading = Units; func main() { _ = Reading(3).sign() }",
    alias_to_defined_struct_inherits_method => "package main; type Row struct { n int }; func (r Row) total() int { return r.n }; type Alias = Row; func main() { _ = Alias{n: 2}.total() }",
    defined_int_pointer_receiver_compile => "package main; type Counter int; func (c *Counter) bump() { *c = *c + 1 }; func main() { value := Counter(0); value.bump(); _ = value }",
}
