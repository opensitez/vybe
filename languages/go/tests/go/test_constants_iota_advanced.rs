//! Advanced `iota` constant blocks: bit-flag OR, offsets, blank skips,
//! typed groups (`int64`, `byte`, `rune`), duplicate values, and arithmetic.
//! Distinct from `test_iota_enumerations.rs` and `test_constants.rs`.

go_run_cases! {
    iota_bit_flags_or_combined => (
        "package main; import \"fmt\"; const ( Read = 1 << iota; Write; Execute ); func main() { fmt.Println(Read | Write | Execute) }",
        vec!["7"]
    ),
    iota_bit_flags_or_pair => (
        "package main; import \"fmt\"; const ( A = 1 << iota; B; C ); func main() { fmt.Println(A | C) }",
        vec!["5"]
    ),
    iota_offset_add_ten => (
        "package main; import \"fmt\"; const ( Base = iota + 10; Next; Last ); func main() { fmt.Println(Base); fmt.Println(Last) }",
        vec!["10", "12"]
    ),
    iota_offset_multiply_three => (
        "package main; import \"fmt\"; const ( Start = iota * 3; Mid; End ); func main() { fmt.Println(Start); fmt.Println(End) }",
        vec!["0", "6"]
    ),
    iota_blank_line_double_skip => (
        "package main; import \"fmt\"; const ( _ = iota; _; X; Y ); func main() { fmt.Println(X); fmt.Println(Y) }",
        vec!["2", "3"]
    ),
    iota_blank_then_explicit => (
        "package main; import \"fmt\"; const ( _ = iota; Z = iota + 5 ); func main() { fmt.Println(Z) }",
        vec!["6"]
    ),
    iota_typed_int64_values => (
        "package main; import \"fmt\"; const ( A int64 = iota; B; C ); func main() { fmt.Println(A); fmt.Println(C) }",
        vec!["0", "2"]
    ),
    iota_typed_byte_sequence => (
        "package main; import \"fmt\"; const ( First byte = iota; Second; Third ); func main() { fmt.Println(int(First)); fmt.Println(int(Third)) }",
        vec!["0", "2"]
    ),
    iota_typed_rune_chars => (
        "package main; import \"fmt\"; const ( Alpha rune = 'A' + iota; Beta; Gamma ); func main() { fmt.Println(int(Alpha)); fmt.Println(int(Gamma)) }",
        vec!["65", "67"]
    ),
    iota_duplicate_values_explicit_repeat => (
        "package main; import \"fmt\"; const ( Low = iota; Mid = Low; High = iota ); func main() { fmt.Println(Low); fmt.Println(Mid); fmt.Println(High) }",
        vec!["0", "0", "2"]
    ),
    iota_arithmetic_square => (
        "package main; import \"fmt\"; const ( A = iota * iota; B; C ); func main() { fmt.Println(A); fmt.Println(B); fmt.Println(C) }",
        vec!["0", "1", "4"]
    ),
    iota_arithmetic_difference => (
        "package main; import \"fmt\"; const ( A = 10 - iota; B; C ); func main() { fmt.Println(A); fmt.Println(C) }",
        vec!["10", "8"]
    ),
    iota_shift_left_per_step => (
        "package main; import \"fmt\"; const ( Bit0 = 1 << iota; Bit1; Bit2; Bit3 ); func main() { fmt.Println(Bit0); fmt.Println(Bit3) }",
        vec!["1", "8"]
    ),
    iota_shift_right_descending => (
        "package main; import \"fmt\"; const ( A = 8 >> iota; B; C ); func main() { fmt.Println(A); fmt.Println(C) }",
        vec!["8", "2"]
    ),
    iota_second_group_restarts => (
        "package main; import \"fmt\"; const ( A = iota; B ); const ( C = iota; D ); func main() { fmt.Println(B); fmt.Println(D) }",
        vec!["1", "1"]
    ),
    iota_mixed_expression_and_implicit => (
        "package main; import \"fmt\"; const ( Seed = iota * 2; Step; Tail ); func main() { fmt.Println(Seed); fmt.Println(Tail) }",
        vec!["0", "4"]
    ),
    iota_negative_values => (
        "package main; import \"fmt\"; const ( A = -iota; B; C ); func main() { fmt.Println(A); fmt.Println(B); fmt.Println(C) }",
        vec!["0", "-1", "-2"]
    ),
    iota_storage_kb_mb_gb => (
        "package main; import \"fmt\"; const ( KB = 1 << (10 * iota); MB; GB ); func main() { fmt.Println(KB); fmt.Println(MB) }",
        vec!["1", "1048576"]
    ),
    iota_bit_and_mask => (
        "package main; import \"fmt\"; const ( M0 = 1 << iota; M1; M2 ); func main() { fmt.Println(M0 & M1); fmt.Println(M0 | M1) }",
        vec!["0", "3"]
    ),
    iota_typed_uint32 => (
        "package main; import \"fmt\"; const ( U0 uint32 = iota; U1; U2 ); func main() { fmt.Println(U0); fmt.Println(U2) }",
        vec!["0", "2"]
    ),
    iota_parenthesized_expression => (
        "package main; import \"fmt\"; const ( V = (iota + 1) * 2; W; X ); func main() { fmt.Println(V); fmt.Println(X) }",
        vec!["2", "6"]
    ),
    iota_combined_with_const_outside => (
        "package main; import \"fmt\"; const offset = 3; const ( A = iota + offset; B ); func main() { fmt.Println(A); fmt.Println(B) }",
        vec!["3", "4"]
    ),
    iota_three_blanks_then_value => (
        "package main; import \"fmt\"; const ( _ = iota; _; _; Target ); func main() { fmt.Println(Target) }",
        vec!["3"]
    ),
    iota_float_conversion => (
        "package main; import \"fmt\"; const ( F0 = float64(iota); F1; F2 ); func main() { fmt.Println(F0); fmt.Println(F2) }",
        vec!["0", "2"]
    ),
    iota_modulo_pattern => (
        "package main; import \"fmt\"; const ( A = iota % 2; B; C; D ); func main() { fmt.Println(A); fmt.Println(B); fmt.Println(C); fmt.Println(D) }",
        vec!["0", "1", "0", "1"]
    ),
    iota_power_of_three => (
        "package main; import \"fmt\"; const ( P0 = 1; P1 = 3 * iota; P2 ); func main() { fmt.Println(P0); fmt.Println(P1); fmt.Println(P2) }",
        vec!["1", "3", "6"]
    ),
    iota_or_chain_three_flags => (
        "package main; import \"fmt\"; const ( F1 = 1 << iota; F2; F3 ); func main() { mask := F1 | F2; fmt.Println(mask); fmt.Println(mask | F3) }",
        vec!["3", "7"]
    ),
    iota_string_number_interleave => (
        "package main; import \"fmt\"; const ( Name = \"v\"; Code = iota; Next ); func main() { fmt.Println(Name); fmt.Println(Code); fmt.Println(Next) }",
        vec!["v", "0", "1"]
    ),
    iota_byte_hex_shift => (
        "package main; import \"fmt\"; const ( H0 byte = 0x10 << iota; H1; H2 ); func main() { fmt.Println(int(H0)); fmt.Println(int(H2)) }",
        vec!["16", "64"]
    ),
    iota_rune_offset_from_a => (
        "package main; import \"fmt\"; const ( R0 rune = 'a' + iota; R1; R2 ); func main() { fmt.Println(string(R0)); fmt.Println(string(R2)) }",
        vec!["a", "c"]
    ),
    iota_duplicate_via_expression => (
        "package main; import \"fmt\"; const ( X = iota; Y = X + 0; Z = iota ); func main() { fmt.Println(X); fmt.Println(Y); fmt.Println(Z) }",
        vec!["0", "0", "2"]
    ),
    iota_int64_large_shift => (
        "package main; import \"fmt\"; const ( T0 int64 = 1 << iota; T1; T2 ); func main() { fmt.Println(T0); fmt.Println(T2) }",
        vec!["1", "4"]
    ),
    iota_subtract_from_base => (
        "package main; import \"fmt\"; const ( A = 5 - iota; B; C ); func main() { fmt.Println(A); fmt.Println(C) }",
        vec!["5", "3"]
    ),
    iota_xor_toggle_bits => (
        "package main; import \"fmt\"; const ( B0 = 1 << iota; B1 ); func main() { fmt.Println(B0 ^ B1) }",
        vec!["3"]
    ),
    iota_reset_in_new_const_block => (
        "package main; import \"fmt\"; const ( A = iota; B = 10 ); const ( C = iota; D ); func main() { fmt.Println(B); fmt.Println(C); fmt.Println(D) }",
        vec!["10", "0", "1"]
    ),
}

go_compile_cases! {
    iota_typed_int64_group_compile =>
        "package main; const ( N0 int64 = iota; N1; N2 ); func main() { _, _, _ = N0, N1, N2 }",
    iota_typed_byte_group_compile =>
        "package main; const ( B0 byte = iota; B1 ); func main() { _, _ = B0, B1 }",
    iota_typed_rune_group_compile =>
        "package main; const ( R0 rune = '!' + iota; R1 ); func main() { _, _ = R0, R1 }",
    iota_blank_multiple_skips_compile =>
        "package main; const ( _ = iota; _; _; V ); func main() { _ = V }",
    iota_bit_or_in_const_expr_compile =>
        "package main; const ( A = 1 << iota; B; Mask = A | B ); func main() { _ = Mask }",
    iota_offset_plus_iota_compile =>
        "package main; const ( Base = 100 + iota; Next ); func main() { _, _ = Base, Next }",
    iota_arithmetic_times_iota_compile =>
        "package main; const ( X = iota * 5; Y; Z ); func main() { _, _, _ = X, Y, Z }",
    iota_duplicate_same_value_compile =>
        "package main; const ( P = iota; Q = P; R = iota ); func main() { _, _, _ = P, Q, R }",
    iota_in_inner_const_block_compile =>
        "package main; const ( outer = 1; inner = iota; tail ); func main() { _, _ = inner, tail }",
    iota_with_explicit_type_int_compile =>
        "package main; const ( A int = iota; B int ); func main() { _, _ = A, B }",
    iota_float64_typed_compile =>
        "package main; const ( F float64 = iota; G ); func main() { _, _ = F, G }",
    iota_complex_expression_compile =>
        "package main; const ( V = (iota + 2) * (iota + 1); W ); func main() { _, _ = V, W }",
    iota_shift_with_iota_multiplier_compile =>
        "package main; const ( S = 1 << (iota + 1); T ); func main() { _, _ = S, T }",
    iota_second_group_after_non_iota_compile =>
        "package main; const X = 9; const ( A = iota; B ); func main() { _, _ = A, B }",
    iota_used_in_var_init_compile =>
        "package main; const ( A = iota; B ); var total = A + B; func main() { _ = total }",
    iota_in_struct_tag_adjacent_compile =>
        "package main; const ( Tag = iota; Other ); type item struct { id int }; func main() { _ = Tag + Other }",
    iota_negative_expression_compile =>
        "package main; const ( N = -1 - iota; M ); func main() { _, _ = N, M }",
    iota_parenthesized_list_compile =>
        "package main; const ( X, Y = iota, iota + 10 ); func main() { _, _ = X, Y }",
    iota_string_const_then_iota_compile =>
        "package main; const ( Label = \"go\"; Code = iota; Next ); func main() { _, _ = Code, Next }",
    iota_uint8_typed_compile =>
        "package main; const ( U uint8 = iota; V ); func main() { _, _ = U, V }",
    iota_or_three_flags_compile =>
        "package main; const ( F0 = 1 << iota; F1; F2; All = F0 | F1 | F2 ); func main() { _ = All }",
    iota_blank_iota_only_compile =>
        "package main; const ( _ = iota; K = iota ); func main() { _ = K }",
    iota_mod_expression_compile =>
        "package main; const ( A = iota % 3; B; C; D ); func main() { _, _, _, _ = A, B, C, D }",
    iota_in_array_length_compile =>
        "package main; const size = iota + 3; const ( A = iota; B ); func main() { arr := [size]int{}; _ = arr; _ = A }",
    iota_combined_add_and_shift_compile =>
        "package main; const ( V = 1<<iota + iota; W ); func main() { _, _ = V, W }",
    iota_three_const_blocks_compile =>
        "package main; const ( A = iota ); const ( B = iota ); const ( C = iota ); func main() { _, _, _ = A, B, C }",
    iota_rune_from_iota_offset_compile =>
        "package main; const ( R rune = '0' + iota; S ); func main() { _, _ = R, S }",
    iota_int32_typed_group_compile =>
        "package main; const ( I int32 = iota; J; K ); func main() { _, _, _ = I, J, K }",
    iota_explicit_value_breaks_chain_compile =>
        "package main; const ( Start = 5; Next = iota; After ); func main() { _, _ = Next, After }",
    iota_in_comparison_switch_compile =>
        "package main; const ( A = iota; B; C ); func main() { switch B { case 1: _ = A + C } }",
}
