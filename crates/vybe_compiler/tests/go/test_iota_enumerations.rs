//! Iota: distinct enumeration patterns (bitmasks, offsets, typed constants).

use crate::helpers::*;

go_run_cases! {
    iota_simple_three_values => ("package main; import \"fmt\"; const ( A = iota; B; C ); func main() { fmt.Println(A); fmt.Println(B); fmt.Println(C) }", vec!["0", "1", "2"]),
    iota_skip_with_blank => ("package main; import \"fmt\"; const ( _ = iota; X = iota; Y ); func main() { fmt.Println(X); fmt.Println(Y) }", vec!["1", "2"]),
    iota_expression_offset => ("package main; import \"fmt\"; const ( A = iota + 10; B; C ); func main() { fmt.Println(A); fmt.Println(C) }", vec!["10", "12"]),
    iota_bitmask_powers => ("package main; import \"fmt\"; const ( FlagA = 1 << iota; FlagB; FlagC ); func main() { fmt.Println(FlagA); fmt.Println(FlagB); fmt.Println(FlagC) }", vec!["1", "2", "4"]),
    iota_per_const_reset => ("package main; import \"fmt\"; const ( A = iota; B ); const ( C = iota; D ); func main() { fmt.Println(B); fmt.Println(D) }", vec!["1", "1"]),
    iota_typed_constants => ("package main; import \"fmt\"; type status int; const ( Ok status = iota; Err; Unknown ); func main() { fmt.Println(int(Ok)); fmt.Println(int(Unknown)) }", vec!["0", "2"]),
    iota_multiplier => ("package main; import \"fmt\"; const ( KB = 1 << (10 * iota); MB ); func main() { fmt.Println(KB); fmt.Println(MB) }", vec!["1024", "1048576"]),
    iota_negative_step => ("package main; import \"fmt\"; const ( A = -iota; B; C ); func main() { fmt.Println(A); fmt.Println(C) }", vec!["0", "-2"]),
}

go_compile_cases! {
    iota_in_parenthesized_group => "package main; const ( X, Y = iota, iota + 1 ); func main() { _, _ = X, Y }",
    iota_with_explicit_value_then_iota => "package main; const ( Start = 5; Next = iota; After ); func main() { _, _ = Next, After }",
    iota_string_constants => "package main; const ( A = \"a\"; B = iota ); func main() { _ = B }",
    iota_float_constant => "package main; const ( F = 1.0 + float64(iota); G ); func main() { _ = G }",
    iota_in_function_scope_compile => "package main; func main() { const X = iota }",
}
