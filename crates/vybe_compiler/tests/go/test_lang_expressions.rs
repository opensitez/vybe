//! Expressions, operators, and evaluation order — one distinct rule per test.

use crate::helpers::*;

go_run_cases! {
    precedence_mul_before_add => ("package main; import \"fmt\"; func main() { fmt.Println(2 + 3 * 4) }", vec!["14"]),
    precedence_paren_override => ("package main; import \"fmt\"; func main() { fmt.Println((2 + 3) * 4) }", vec!["20"]),
    short_circuit_and_false => ("package main; import \"fmt\"; func main() { fmt.Println(false && (1/0 == 1)) }", vec!["false"]),
    short_circuit_or_true => ("package main; import \"fmt\"; func main() { fmt.Println(true || (1/0 == 1)) }", vec!["true"]),
    bitwise_and_or => ("package main; import \"fmt\"; func main() { fmt.Println(5 & 3); fmt.Println(5 | 1) }", vec!["1", "5"]),
    bitwise_xor_shift => ("package main; import \"fmt\"; func main() { fmt.Println(10 ^ 12); fmt.Println(1 << 3) }", vec!["6", "8"]),
    unary_plus_minus => ("package main; import \"fmt\"; func main() { x := 5; fmt.Println(-x); fmt.Println(+x) }", vec!["-5", "5"]),
    comparison_chained_false => ("package main; import \"fmt\"; func main() { fmt.Println(1 < 2 && 2 < 3) }", vec!["true"]),
    string_concat => ("package main; import \"fmt\"; func main() { fmt.Println(\"a\" + \"b\") }", vec!["ab"]),
    numeric_conversion_in_expr => ("package main; import \"fmt\"; func main() { var b byte = 65; fmt.Println(string(b)) }", vec!["A"]),
    address_of_composite_element => ("package main; import \"fmt\"; func main() { s := []int{1}; p := &s[0]; *p = 9; fmt.Println(s[0]) }", vec!["9"]),
    dereference_pointer => ("package main; import \"fmt\"; func main() { x := 2; p := &x; *p = 7; fmt.Println(x) }", vec!["7"]),
    struct_pointer_field_arrow => ("package main; import \"fmt\"; type P struct { N int }; func main() { p := &P{N:1}; p.N = 2; fmt.Println(p.N) }", vec!["2"]),
    index_mutable_slice => ("package main; import \"fmt\"; func main() { s := []int{1,2,3}; s[1] = 9; fmt.Println(s[1]) }", vec!["9"]),
    map_assignment_insert => ("package main; import \"fmt\"; func main() { m := map[string]int{}; m[\"k\"] = 4; fmt.Println(m[\"k\"]) }", vec!["4"]),
    increment_statement => ("package main; import \"fmt\"; func main() { n := 1; n++; fmt.Println(n) }", vec!["2"]),
    decrement_statement => ("package main; import \"fmt\"; func main() { n := 2; n--; fmt.Println(n) }", vec!["1"]),
    compound_add_assign => ("package main; import \"fmt\"; func main() { n := 1; n += 3; fmt.Println(n) }", vec!["4"]),
    compound_bit_and_assign => ("package main; import \"fmt\"; func main() { n := 7; n &= 3; fmt.Println(n) }", vec!["3"]),
    comma_operator_not_exists_use_multi_assign => ("package main; import \"fmt\"; func main() { a, b := 1, 2; fmt.Println(a+b) }", vec!["3"]),
    blank_in_range => ("package main; import \"fmt\"; func main() { sum := 0; for _, v := range []int{1,2} { sum += v }; fmt.Println(sum) }", vec!["3"]),
    for_three_clause => ("package main; import \"fmt\"; func main() { for i := 0; i < 2; i++ { if i == 1 { fmt.Println(i) } } }", vec!["1"]),
    for_range_slice_value_only => ("package main; import \"fmt\"; func main() { for _, v := range []int{4} { fmt.Println(v) } }", vec!["4"]),
    switch_expression_no_init => ("package main; import \"fmt\"; func main() { switch 2 { case 2: fmt.Println(\"y\") } }", vec!["y"]),
    switch_tagless_true => ("package main; import \"fmt\"; func main() { switch { case 1 < 2: fmt.Println(\"t\") } }", vec!["t"]),
    type_assertion_single_value => ("package main; import \"fmt\"; func main() { var i interface{} = 1; fmt.Println(i.(int)) }", vec!["1"]),
    len_array => ("package main; import \"fmt\"; func main() { fmt.Println(len([3]int{})) }", vec!["3"]),
    cap_slice => ("package main; import \"fmt\"; func main() { s := make([]int, 2, 5); fmt.Println(cap(s)) }", vec!["5"]),
    append_returns_same_slice_header => ("package main; import \"fmt\"; func main() { s := []int{1}; t := append(s, 2); fmt.Println(t[1]) }", vec!["2"]),
    new_returns_pointer => ("package main; import \"fmt\"; func main() { p := new(int); *p = 6; fmt.Println(*p) }", vec!["6"]),
    make_slice_len => ("package main; import \"fmt\"; func main() { s := make([]int, 3); fmt.Println(len(s)) }", vec!["3"]),
    make_map_ready => ("package main; import \"fmt\"; func main() { m := make(map[string]int); m[\"a\"] = 1; fmt.Println(m[\"a\"]) }", vec!["1"]),
    make_chan_buffered => ("package main; import \"fmt\"; func main() { ch := make(chan int, 2); fmt.Println(cap(ch)) }", vec!["2"]),
    function_call_args_eval_left_to_right => ("package main; import \"fmt\"; func f(a, b int) int { return a*10+b }; func main() { i := 0; i++; fmt.Println(f(i, i)) }", vec!["12"]),
    method_call_on_literal => ("package main; import \"fmt\"; func main() { s := []int{1,2}; fmt.Println(len(s)) }", vec!["2"]),
}

go_compile_cases! {
    invalid_div_zero_compile_still_parse => "package main; func main() { _ = 1 / 0 }",
    shift_negative_compile => "package main; func main() { var n int; _ = n >> -1 }",
    take_address_rvalue_compile => "package main; func main() { _ = &([]int{1}[0]) }",
}
