use super::helpers::run_prints;

#[test]
fn array_constructor_repetition_counts_basic_pairing() {
    let out = run_prints(
        r#"
program array_constructor_repetition_counts_basic_pairing
    integer, allocatable :: values(:)
    values = (/ 2 * 10, 3 * 20 /)
    print *, size(values)
    print *, sum(values)
    print *, values(1)
    print *, values(size(values))
end program array_constructor_repetition_counts_basic_pairing
"#,
    );
    assert_eq!(out, vec!["5", "80", "10", "20"]);
}

#[test]
fn array_constructor_repetition_counts_negative_repeated_term() {
    let out = run_prints(
        r#"
program array_constructor_repetition_counts_negative_repeated_term
    integer, allocatable :: values(:)
    values = (/ 3 * -4, 2 * 5 /)
    print *, size(values)
    print *, sum(values)
    print *, values(1)
    print *, values(size(values))
end program array_constructor_repetition_counts_negative_repeated_term
"#,
    );
    assert_eq!(out, vec!["5", "-2", "-4", "5"]);
}

#[test]
fn array_constructor_repetition_counts_variable_expression_counts() {
    let out = run_prints(
        r#"
program array_constructor_repetition_counts_variable_expression_counts
    integer, allocatable :: values(:)
    values = (/ (2 + 1) * 4, (1 + 2) * 1, (2 + 1) * -3 /)
    print *, size(values)
    print *, sum(values)
    print *, values(1)
    print *, values(size(values))
end program array_constructor_repetition_counts_variable_expression_counts
"#,
    );
    assert_eq!(out, vec!["5", "1", "4", "-3"]);
}

#[test]
fn array_constructor_repetition_counts_zero_value_blocks() {
    let out = run_prints(
        r#"
program array_constructor_repetition_counts_zero_value_blocks
    integer, allocatable :: values(:)
    values = (/ 3 * 0, 2 * 7, 1 * 0 /)
    print *, size(values)
    print *, sum(values)
    print *, values(1)
    print *, values(size(values))
end program array_constructor_repetition_counts_zero_value_blocks
"#,
    );
    assert_eq!(out, vec!["6", "14", "0", "0"]);
}

#[test]
fn array_constructor_repetition_counts_singleton_implicit_default() {
    let out = run_prints(
        r#"
program array_constructor_repetition_counts_singleton_implicit_default
    integer, allocatable :: values(:)
    values = (/ 6 * 1, 4, 5 * 2 /)
    print *, size(values)
    print *, sum(values)
    print *, values(1)
    print *, values(size(values))
end program array_constructor_repetition_counts_singleton_implicit_default
"#,
    );
    assert_eq!(out, vec!["11", "24", "1", "2"]);
}

#[test]
fn array_constructor_repetition_counts_descending_repeat_pattern() {
    let out = run_prints(
        r#"
program array_constructor_repetition_counts_descending_repeat_pattern
    integer, allocatable :: values(:)
    values = (/ 3 * 9, 1 * 8, 2 * 7, 1 * 6, 1 /)
    print *, size(values)
    print *, sum(values)
    print *, values(1)
    print *, values(size(values))
end program array_constructor_repetition_counts_descending_repeat_pattern
"#,
    );
    assert_eq!(out, vec!["8", "55", "9", "1"]);
}

#[test]
fn array_constructor_repetition_counts_mixed_repeats_and_implied_do_tail() {
    let out = run_prints(
        r#"
program array_constructor_repetition_counts_mixed_repeats_and_implied_do_tail
    integer, allocatable :: values(:)
    values = (/ 3 * 2, (i, i = 1, 3), 1 * 12 /)
    print *, size(values)
    print *, sum(values)
    print *, values(1)
    print *, values(size(values))
    print *, values(4)
    print *, values(5)
end program array_constructor_repetition_counts_mixed_repeats_and_implied_do_tail
"#,
    );
    assert_eq!(out, vec!["7", "24", "2", "12", "1", "2"]);
}

#[test]
fn array_constructor_repetition_counts_implied_do_prefix_then_repeats() {
    let out = run_prints(
        r#"
program array_constructor_repetition_counts_implied_do_prefix_then_repeats
    integer, allocatable :: values(:)
    values = (/ (i, i = 1, 4), 2 * 5 /)
    print *, size(values)
    print *, sum(values)
    print *, values(1)
    print *, values(size(values))
end program array_constructor_repetition_counts_implied_do_prefix_then_repeats
"#,
    );
    assert_eq!(out, vec!["6", "29", "1", "5"]);
}

#[test]
fn array_constructor_repetition_counts_implied_do_and_nested_repeats() {
    let out = run_prints(
        r#"
program array_constructor_repetition_counts_implied_do_and_nested_repeats
    integer, allocatable :: values(:)
    values = (/ (i, i = 2, 4), 2 * 10, 1 * (3 + 2), 1 * 0 /)
    print *, size(values)
    print *, sum(values)
    print *, values(1)
    print *, values(6)
end program array_constructor_repetition_counts_implied_do_and_nested_repeats
"#,
    );
    assert_eq!(out, vec!["8", "33", "2", "0"]);
}

#[test]
fn array_constructor_repetition_counts_fixed_shape_integer_array() {
    let out = run_prints(
        r#"
program array_constructor_repetition_counts_fixed_shape_integer_array
    integer :: values(6)
    values = (/ 3 * 4, 2 * 6, 1 * 10 /)
    print *, size(values)
    print *, values(1)
    print *, values(4)
    print *, sum(values)
end program array_constructor_repetition_counts_fixed_shape_integer_array
"#,
    );
    assert_eq!(out, vec!["6", "4", "6", "34"]);
}

#[test]
fn array_constructor_repetition_counts_fixed_shape_real_array() {
    let out = run_prints(
        r#"
program array_constructor_repetition_counts_fixed_shape_real_array
    real :: values(5)
    integer :: n
    values = (/ 2 * 1.25, 3 * 0.75 /)
    n = size(values)
    print *, n
    print *, nint(sum(values))
    print *, nint(values(1))
    print *, nint(values(n))
end program array_constructor_repetition_counts_fixed_shape_real_array
"#,
    );
    assert_eq!(out, vec!["5", "5", "1", "0"]);
}

#[test]
fn array_constructor_repetition_counts_large_block_explicit() {
    let out = run_prints(
        r#"
program array_constructor_repetition_counts_large_block_explicit
    integer, allocatable :: values(:)
    values = (/ 8 * 1, 1, 1 * -2, 2 * 2 /)
    print *, size(values)
    print *, sum(values)
    print *, values(1)
    print *, values(size(values))
end program array_constructor_repetition_counts_large_block_explicit
"#,
    );
    assert_eq!(out, vec!["12", "6", "1", "2"]);
}

#[test]
fn array_constructor_repetition_counts_chain_of_single_repeats() {
    let out = run_prints(
        r#"
program array_constructor_repetition_counts_chain_of_single_repeats
    integer, allocatable :: values(:)
    values = (/ 1 * 8, 1 * 1, 1 * 6, 1 * 4 /)
    print *, size(values)
    print *, sum(values)
    print *, values(1)
    print *, values(size(values))
end program array_constructor_repetition_counts_chain_of_single_repeats
"#,
    );
    assert_eq!(out, vec!["4", "19", "8", "4"]);
}

#[test]
fn array_constructor_repetition_counts_repeat_of_expression_results() {
    let out = run_prints(
        r#"
program array_constructor_repetition_counts_repeat_of_expression_results
    integer, allocatable :: values(:)
    values = (/ 3 * (1 + 2), 2 * (5 - 3) /)
    print *, size(values)
    print *, sum(values)
    print *, values(1)
    print *, values(size(values))
end program array_constructor_repetition_counts_repeat_of_expression_results
"#,
    );
    assert_eq!(out, vec!["5", "13", "3", "2"]);
}

#[test]
fn array_constructor_repetition_counts_repeat_with_sign() {
    let out = run_prints(
        r#"
program array_constructor_repetition_counts_repeat_with_sign
    integer, allocatable :: values(:)
    values = (/ 4 * (-1), 1 * 6, 1 * (-3) /)
    print *, size(values)
    print *, sum(values)
    print *, values(1)
    print *, values(size(values))
end program array_constructor_repetition_counts_repeat_with_sign
"#,
    );
    assert_eq!(out, vec!["6", "3", "-1", "-3"]);
}

#[test]
fn array_constructor_repetition_counts_repeat_and_division_tail() {
    let out = run_prints(
        r#"
program array_constructor_repetition_counts_repeat_and_division_tail
    integer, allocatable :: values(:)
    values = (/ 2 * 12, 1 * (20 / 4), 2 * 2 /)
    print *, size(values)
    print *, sum(values)
    print *, values(1)
    print *, values(size(values))
end program array_constructor_repetition_counts_repeat_and_division_tail
"#,
    );
    assert_eq!(out, vec!["5", "34", "12", "2"]);
}

#[test]
fn array_constructor_repetition_counts_repeat_with_reduction_check() {
    let out = run_prints(
        r#"
program array_constructor_repetition_counts_repeat_with_reduction_check
    integer, allocatable :: values(:)
    values = (/ 5 * 1, 3 * 2, 2 * 3 /)
    print *, size(values)
    print *, sum(values)
    print *, count(values == 2)
    print *, values(1)
end program array_constructor_repetition_counts_repeat_with_reduction_check
"#,
    );
    assert_eq!(out, vec!["10", "19", "3", "1"]);
}

#[test]
fn array_constructor_repetition_counts_nested_constructor_roundtrip() {
    let out = run_prints(
        r#"
program array_constructor_repetition_counts_nested_constructor_roundtrip
    integer, allocatable :: values(:)
    values = (/ 2 * (3 * 1), 1 * (2 * 2), 1 * (2 + 1) /)
    print *, size(values)
    print *, sum(values)
    print *, values(2)
    print *, values(size(values))
end program array_constructor_repetition_counts_nested_constructor_roundtrip
"#,
    );
    assert_eq!(out, vec!["4", "9", "3", "3"]);
}

#[test]
fn array_constructor_repetition_counts_repeated_zero_length_while_not_empty() {
    let out = run_prints(
        r#"
program array_constructor_repetition_counts_repeated_zero_length_while_not_empty
    integer, allocatable :: values(:)
    values = (/ 4 * 0, 2 * 9 /)
    print *, size(values)
    print *, sum(values)
    print *, values(1)
    print *, values(size(values))
end program array_constructor_repetition_counts_repeated_zero_length_while_not_empty
"#,
    );
    assert_eq!(out, vec!["6", "18", "0", "9"]);
}

#[test]
fn array_constructor_repetition_counts_even_odd_parity_mix() {
    let out = run_prints(
        r#"
program array_constructor_repetition_counts_even_odd_parity_mix
    integer, allocatable :: values(:)
    values = (/ 3 * 2, 3 * 3, 2 * 2 /)
    print *, size(values)
    print *, sum(values)
    print *, values(1)
    print *, values(4)
    print *, values(size(values))
end program array_constructor_repetition_counts_even_odd_parity_mix
"#,
    );
    assert_eq!(out, vec!["8", "18", "2", "2", "2"]);
}

#[test]
fn array_constructor_repetition_counts_character_repeat_vector() {
    let out = run_prints(
        r#"
program array_constructor_repetition_counts_character_repeat_vector
    character(len=4), allocatable :: values(:)
    values = (/ 3 * 'a', 2 * 'xy' /)
    print *, size(values)
    print *, len(values(1))
    print *, len_trim(values(2))
    print *, len_trim(values(4))
end program array_constructor_repetition_counts_character_repeat_vector
"#,
    );
    assert_eq!(out, vec!["5", "4", "1", "2"]);
}

#[test]
fn array_constructor_repetition_counts_logical_repetition_through_merge() {
    let out = run_prints(
        r#"
program array_constructor_repetition_counts_logical_repetition_through_merge
    logical, allocatable :: flags(:)
    integer :: n
    flags = (/ 2 * .true., 3 * .false. /)
    n = size(flags)
    print *, n
    print *, count(flags)
    print *, merge(1, 0, flags(1))
    print *, merge(1, 0, flags(n))
end program array_constructor_repetition_counts_logical_repetition_through_merge
"#,
    );
    assert_eq!(out, vec!["5", "2", "1", "0"]);
}
