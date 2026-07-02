//! cmp and constraints (Go 1.21): Compare, Less, Or, Clamp patterns.


go_run_cases! {
    cmp_compare_int_less => ("package main; import \"fmt\"; import \"cmp\"; func main() { fmt.Println(cmp.Compare(1, 2)) }", vec!["-1"]),
    cmp_compare_int_equal => ("package main; import \"fmt\"; import \"cmp\"; func main() { fmt.Println(cmp.Compare(3, 3)) }", vec!["0"]),
    cmp_compare_int_greater => ("package main; import \"fmt\"; import \"cmp\"; func main() { fmt.Println(cmp.Compare(5, 2)) }", vec!["1"]),
    cmp_less_string => ("package main; import \"fmt\"; import \"cmp\"; func main() { fmt.Println(cmp.Less(\"a\", \"b\")) }", vec!["true"]),
    cmp_or_first_nonzero => ("package main; import \"fmt\"; import \"cmp\"; func main() { fmt.Println(cmp.Or(0, 0, 7)) }", vec!["7"]),
    cmp_clamp_within => ("package main; import \"fmt\"; import \"cmp\"; func main() { fmt.Println(cmp.Clamp(5, 1, 10)) }", vec!["5"]),
    cmp_clamp_below_lo => ("package main; import \"fmt\"; import \"cmp\"; func main() { fmt.Println(cmp.Clamp(-1, 0, 9)) }", vec!["0"]),
    cmp_clamp_above_hi => ("package main; import \"fmt\"; import \"cmp\"; func main() { fmt.Println(cmp.Clamp(99, 0, 9)) }", vec!["9"]),
}

go_compile_cases! {
    cmp_compare_float64 => "package main; import \"cmp\"; func main() { _ = cmp.Compare(1.5, 2.0) }",
    cmp_less_float64 => "package main; import \"cmp\"; func main() { _ = cmp.Less(1.0, 2.0) }",
    cmp_or_strings => "package main; import \"cmp\"; func main() { _ = cmp.Or(\"\", \"ok\") }",
}
