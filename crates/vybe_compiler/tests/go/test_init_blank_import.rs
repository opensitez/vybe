//! package init() ordering, side-effect imports, and blank imports (`_ "pkg"`).


go_run_cases! {
    init_order_three_sequential_appends => (
        "package main; import \"fmt\"; var order string; func init() { order = order + \"1\" }; func init() { order = order + \"2\" }; func init() { order = order + \"3\" }; func main() { fmt.Println(order) }",
        vec!["123"]
    ),
    init_order_numeric_accumulation => (
        "package main; import \"fmt\"; var total int; func init() { total += 1 }; func init() { total += 10 }; func init() { total += 100 }; func main() { fmt.Println(total) }",
        vec!["111"]
    ),
    init_populates_package_slice_before_main => (
        "package main; import \"fmt\"; var values []int; func init() { values = append(values, 2, 4) }; func main() { fmt.Println(len(values)); fmt.Println(values[1]) }",
        vec!["2", "4"]
    ),
    init_fills_map_lookup_before_main => (
        "package main; import \"fmt\"; var table = map[string]int{}; func init() { table[\"go\"] = 7 }; func main() { fmt.Println(table[\"go\"]) }",
        vec!["7"]
    ),
    init_calls_package_helper_before_main => (
        "package main; import \"fmt\"; var ready bool; func markReady() { ready = true }; func init() { markReady() }; func main() { fmt.Println(ready) }",
        vec!["true"]
    ),
}

go_compile_cases! {
    four_init_functions_sequential_compile =>
        "package main; var step int; func init() { step = 1 }; func init() { step++ }; func init() { step++ }; func init() { step++ }; func main() { _ = step }",
    init_with_if_else_branch_compile =>
        "package main; var flag bool; func init() { if 2 > 1 { flag = true } else { flag = false } }; func main() { _ = flag }",
    init_with_for_loop_compile =>
        "package main; var count int; func init() { for i := 0; i < 3; i++ { count++ } }; func main() { _ = count }",
    init_assigns_struct_literal_field_compile =>
        "package main; type node struct { value int }; var root node; func init() { root = node{value: 9} }; func main() { _ = root.value }",
    init_uses_regular_import_in_body_compile =>
        "package main; import \"fmt\"; var label string; func init() { label = fmt.Sprintf(\"%s-%d\", \"init\", 3) }; func main() { _ = label }",
    init_sets_func_variable_compile =>
        "package main; var twice func(int) int; func init() { twice = func(n int) int { return n * 2 } }; func main() { _ = twice(4) }",
    init_with_nested_block_scope_compile =>
        "package main; var depth int; func init() { { depth = 1; { depth = depth + 1 } } }; func main() { _ = depth }",
    init_after_type_declaration_compile =>
        "package main; type score int; var high score; func init() { high = score(99) }; func main() { _ = high }",
    init_mutates_existing_slice_element_compile =>
        "package main; var nums = []int{1, 2, 3}; func init() { nums[1] = 20 }; func main() { _ = nums[1] }",
    init_writes_const_derived_package_var_compile =>
        "package main; const base = 4; var doubled int; func init() { doubled = base * 2 }; func main() { _ = doubled }",
    blank_import_strings_compile =>
        "package main; import _ \"strings\"; func main() {}",
    blank_import_math_compile =>
        "package main; import _ \"math\"; func main() {}",
    grouped_blank_imports_compile =>
        "package main; import ( _ \"strings\"; _ \"math\" ); func main() {}",
    blank_import_mixed_with_named_import_compile =>
        "package main; import ( \"fmt\"; _ \"strings\" ); func main() { _ = fmt.Sprintf(\"%d\", 1) }",
    multiple_blank_imports_different_packages_compile =>
        "package main; import _ \"strings\"; import _ \"math\"; func main() {}",
    blank_import_with_alias_import_compile =>
        "package main; import f \"fmt\"; import _ \"strings\"; func main() { _ = f.Sprint(\"ok\") }",
    blank_import_encoding_json_compile =>
        "package main; import _ \"encoding/json\"; func main() {}",
    init_before_main_with_blank_import_compile =>
        "package main; import _ \"strings\"; var seeded int; func init() { seeded = 5 }; func main() { _ = seeded }",
    init_chain_reads_prior_init_value_compile =>
        "package main; var first int; var second int; func init() { first = 3 }; func init() { second = first + 2 }; func main() { _ = second }",
    init_with_switch_statement_compile =>
        "package main; var tag string; func init() { switch 2 { case 2: tag = \"hit\" default: tag = \"miss\" } }; func main() { _ = tag }",
}
