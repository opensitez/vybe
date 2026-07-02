//! Type switch and switch initialization — distinct control-flow shapes.


go_run_cases! {
    type_switch_int_branch => ("package main; import \"fmt\"; func describe(v interface{}) { switch v.(type) { case int: fmt.Println(\"int\") default: fmt.Println(\"other\") } }; func main() { describe(3) }", vec!["int"]),
    type_switch_string_branch => ("package main; import \"fmt\"; func describe(v interface{}) { switch v.(type) { case string: fmt.Println(\"str\") default: fmt.Println(\"other\") } }; func main() { describe(\"x\") }", vec!["str"]),
    type_switch_default => ("package main; import \"fmt\"; func describe(v interface{}) { switch v.(type) { case int: fmt.Println(\"int\") default: fmt.Println(\"default\") } }; func main() { describe(1.5) }", vec!["default"]),
    type_switch_multi_value => ("package main; import \"fmt\"; func describe(v interface{}) { switch v.(type) { case int, int64: fmt.Println(\"integer\") default: fmt.Println(\"other\") } }; func main() { describe(int64(9)) }", vec!["integer"]),
    switch_init_statement => ("package main; import \"fmt\"; func main() { switch x := 3; x { case 1: fmt.Println(\"one\") case 3: fmt.Println(\"three\") default: fmt.Println(\"other\") } }", vec!["three"]),
    switch_tagless_true => ("package main; import \"fmt\"; func main() { x := 5; switch { case x < 3: fmt.Println(\"low\") case x < 10: fmt.Println(\"mid\") default: fmt.Println(\"high\") } }", vec!["mid"]),
    switch_fallthrough => ("package main; import \"fmt\"; func main() { x := 1; switch x { case 1: fmt.Println(\"a\"); fallthrough; case 2: fmt.Println(\"b\") } }", vec!["a", "b"]),
    switch_expression_go18 => ("package main; import \"fmt\"; func main() { x := 2; switch x { case 1, 2: fmt.Println(\"pair\") default: fmt.Println(\"other\") } }", vec!["pair"]),
}

go_compile_cases! {
    type_switch_bind_variable => "package main; func describe(v interface{}) { switch value := v.(type) { case int: _ = value; default: _ = value } }; func main() { describe(1) }",
    type_switch_interface_case => "package main; type P interface { M() }; func describe(v interface{}) { switch v.(type) { case P: _ = v } }; func main() {}",
    switch_empty_body => "package main; func main() { switch 1 { case 1: } }",
    switch_duplicate_case_compile => "package main; func main() { switch 1 { case 1, 1: } }",
}
