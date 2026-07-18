use super::helpers::run_prints;

#[test]
fn array_constructor_shape_inference_01_direct_literal_values() {
    let out = run_prints(
        r#"
program test_array_constructor_shape_inference_01_direct_literal_values
    integer, allocatable :: values(:)
    values = (/ 1, 2, 3 /)
    print *, size(values)
    print *, sum(values)
    print *, values(1)
    print *, values(size(values))
end program test_array_constructor_shape_inference_01_direct_literal_values
"#,
    );
    assert_eq!(out, vec!["3", "6", "1", "3"]);
}

#[test]
fn array_constructor_shape_inference_02_direct_literal_with_negatives() {
    let out = run_prints(
        r#"
program test_array_constructor_shape_inference_02_direct_literal_with_negatives
    integer, allocatable :: values(:)
    values = (/ -3, 5, -1, 7 /)
    print *, size(values)
    print *, sum(values)
    print *, values(1)
    print *, values(size(values))
end program test_array_constructor_shape_inference_02_direct_literal_with_negatives
"#,
    );
    assert_eq!(out, vec!["4", "8", "-3", "7"]);
}

#[test]
fn array_constructor_shape_inference_03_implied_do_linear_sequence() {
    let out = run_prints(
        r#"
program test_array_constructor_shape_inference_03_implied_do_linear_sequence
    integer, allocatable :: values(:)
    values = (/ (i, i = 1, 5) /)
    print *, size(values)
    print *, sum(values)
    print *, values(1)
    print *, values(size(values))
end program test_array_constructor_shape_inference_03_implied_do_linear_sequence
"#,
    );
    assert_eq!(out, vec!["5", "15", "1", "5"]);
}

#[test]
fn array_constructor_shape_inference_04_implied_do_even_stride() {
    let out = run_prints(
        r#"
program test_array_constructor_shape_inference_04_implied_do_even_stride
    integer, allocatable :: values(:)
    values = (/ (i, i = 2, 10, 2) /)
    print *, size(values)
    print *, sum(values)
    print *, values(1)
    print *, values(size(values))
end program test_array_constructor_shape_inference_04_implied_do_even_stride
"#,
    );
    assert_eq!(out, vec!["5", "30", "2", "10"]);
}

#[test]
fn array_constructor_shape_inference_05_implied_do_descending_stride() {
    let out = run_prints(
        r#"
program test_array_constructor_shape_inference_05_implied_do_descending_stride
    integer, allocatable :: values(:)
    values = (/ (i, i = 11, 3, -3) /)
    print *, size(values)
    print *, sum(values)
    print *, values(1)
    print *, values(size(values))
end program test_array_constructor_shape_inference_05_implied_do_descending_stride
"#,
    );
    assert_eq!(out, vec!["3", "17", "11", "5"]);
}

#[test]
fn array_constructor_shape_inference_06_implied_do_from_variables() {
    let out = run_prints(
        r#"
program test_array_constructor_shape_inference_06_implied_do_from_variables
    integer :: start_idx
    integer :: stop_idx
    integer :: step
    integer, allocatable :: values(:)
    start_idx = 4
    stop_idx = 12
    step = 2
    values = (/ (i, i = start_idx, stop_idx, step) /)
    print *, size(values)
    print *, sum(values)
    print *, values(1)
    print *, values(size(values))
end program test_array_constructor_shape_inference_06_implied_do_from_variables
"#,
    );
    assert_eq!(out, vec!["5", "40", "4", "12"]);
}

#[test]
fn array_constructor_shape_inference_07_implied_do_expression_series() {
    let out = run_prints(
        r#"
program test_array_constructor_shape_inference_07_implied_do_expression_series
    integer, allocatable :: values(:)
    values = (/ (i * 3 - 1, i = 1, 5) /)
    print *, size(values)
    print *, sum(values)
    print *, values(1)
    print *, values(size(values))
end program test_array_constructor_shape_inference_07_implied_do_expression_series
"#,
    );
    assert_eq!(out, vec!["5", "45", "2", "14"]);
}

#[test]
fn array_constructor_shape_inference_08_implied_do_square_series() {
    let out = run_prints(
        r#"
program test_array_constructor_shape_inference_08_implied_do_square_series
    integer, allocatable :: values(:)
    values = (/ (i * i, i = 1, 4) /)
    print *, size(values)
    print *, sum(values)
    print *, values(1)
    print *, values(size(values))
end program test_array_constructor_shape_inference_08_implied_do_square_series
"#,
    );
    assert_eq!(out, vec!["4", "30", "1", "16"]);
}

#[test]
fn array_constructor_shape_inference_09_repetition_single_block() {
    let out = run_prints(
        r#"
program test_array_constructor_shape_inference_09_repetition_single_block
    integer, allocatable :: values(:)
    values = (/ 4 * 7 /)
    print *, size(values)
    print *, sum(values)
    print *, values(1)
    print *, values(size(values))
end program test_array_constructor_shape_inference_09_repetition_single_block
"#,
    );
    assert_eq!(out, vec!["4", "28", "7", "7"]);
}

#[test]
fn array_constructor_shape_inference_10_repetition_mixed_blocks() {
    let out = run_prints(
        r#"
program test_array_constructor_shape_inference_10_repetition_mixed_blocks
    integer, allocatable :: values(:)
    values = (/ 2 * 3, 3 * 8 /)
    print *, size(values)
    print *, sum(values)
    print *, values(1)
    print *, values(size(values))
end program test_array_constructor_shape_inference_10_repetition_mixed_blocks
"#,
    );
    assert_eq!(out, vec!["5", "34", "3", "8"]);
}

#[test]
fn array_constructor_shape_inference_11_repetition_and_implied_do_mix() {
    let out = run_prints(
        r#"
program test_array_constructor_shape_inference_11_repetition_and_implied_do_mix
    integer, allocatable :: values(:)
    values = (/ 2 * 1, (i, i = 2, 4), 3 * 0 /)
    print *, size(values)
    print *, sum(values)
    print *, values(1)
    print *, values(size(values))
end program test_array_constructor_shape_inference_11_repetition_and_implied_do_mix
"#,
    );
    assert_eq!(out, vec!["8", "10", "1", "0"]);
}

#[test]
fn array_constructor_shape_inference_12_allocatable_resize_grows() {
    let out = run_prints(
        r#"
program test_array_constructor_shape_inference_12_allocatable_resize_grows
    integer, allocatable :: values(:)
    values = (/ 11, 22 /)
    print *, size(values)
    values = (/ 1, 2, 3, 4, 5, 6 /)
    print *, size(values)
    print *, sum(values)
    print *, values(size(values))
end program test_array_constructor_shape_inference_12_allocatable_resize_grows
"#,
    );
    assert_eq!(out, vec!["2", "6", "21", "6"]);
}

#[test]
fn array_constructor_shape_inference_13_allocatable_resize_shrinks() {
    let out = run_prints(
        r#"
program test_array_constructor_shape_inference_13_allocatable_resize_shrinks
    integer, allocatable :: values(:)
    values = (/ 1, 2, 3, 4, 5, 6 /)
    print *, size(values)
    values = (/ 99 /)
    print *, size(values)
    print *, sum(values)
    print *, values(1)
end program test_array_constructor_shape_inference_13_allocatable_resize_shrinks
"#,
    );
    assert_eq!(out, vec!["6", "1", "99", "99"]);
}

#[test]
fn array_constructor_shape_inference_14_fixed_shape_initializer() {
    let out = run_prints(
        r#"
program test_array_constructor_shape_inference_14_fixed_shape_initializer
    integer :: values(4) = (/ 2, 4, 6, 8 /)
    print *, size(values)
    print *, sum(values)
    print *, values(1)
    print *, values(size(values))
end program test_array_constructor_shape_inference_14_fixed_shape_initializer
"#,
    );
    assert_eq!(out, vec!["4", "20", "2", "8"]);
}

#[test]
fn array_constructor_shape_inference_15_real_constructor_to_real_alloc() {
    let out = run_prints(
        r#"
program test_array_constructor_shape_inference_15_real_constructor_to_real_alloc
    real, allocatable :: values(:)
    integer :: n
    values = (/ 1.5, 2.5, 3.5 /)
    n = size(values)
    print *, n
    print *, nint(sum(values))
    print *, values(1)
    print *, values(n)
end program test_array_constructor_shape_inference_15_real_constructor_to_real_alloc
"#,
    );
    assert_eq!(out, vec!["3", "7", "1", "3"]);
}

#[test]
fn array_constructor_shape_inference_16_real_implied_do_with_scaling() {
    let out = run_prints(
        r#"
program test_array_constructor_shape_inference_16_real_implied_do_with_scaling
    real, allocatable :: values(:)
    integer :: n
    values = (/ (real(i) * 0.5, i = 2, 8, 2) /)
    n = size(values)
    print *, n
    print *, nint(sum(values))
    print *, values(1)
    print *, values(n)
end program test_array_constructor_shape_inference_16_real_implied_do_with_scaling
"#,
    );
    assert_eq!(out, vec!["4", "10", "1", "4"]);
}

#[test]
fn array_constructor_shape_inference_17_constructor_as_subroutine_argument_size_and_sum() {
    let out = run_prints(
        r#"
program test_array_constructor_shape_inference_17_constructor_as_subroutine_argument_size_and_sum
    call validate_values((/ 3, 1, 4, 1, 5, 9 /), 6, 23, 3)
contains
    subroutine validate_values(values, expected_size, expected_sum, expected_first)
        integer, intent(in) :: values(:)
        integer, intent(in) :: expected_size
        integer, intent(in) :: expected_sum
        integer, intent(in) :: expected_first
        print *, merge(1, 0, size(values) == expected_size)
        print *, merge(1, 0, sum(values) == expected_sum)
        print *, merge(1, 0, values(1) == expected_first)
    end subroutine validate_values
end program test_array_constructor_shape_inference_17_constructor_as_subroutine_argument_size_and_sum
"#,
    );
    assert_eq!(out, vec!["1", "1", "1"]);
}

#[test]
fn array_constructor_shape_inference_18_constructor_assisted_reduction_guard_1() {
    let out = run_prints(
        r#"
program test_array_constructor_shape_inference_18_constructor_assisted_reduction_guard_1
    call validate_values((/ 10, 20, 30, 40, 50 /), 50, 150)
contains
    subroutine validate_values(values, expected_last, expected_sum)
        integer, intent(in) :: values(:)
        integer, intent(in) :: expected_last
        integer, intent(in) :: expected_sum
        integer :: last_value
        last_value = values(size(values))
        print *, merge(1, 0, size(values) >= 3)
        print *, merge(1, 0, sum(values) == expected_sum)
        print *, merge(1, 0, last_value == expected_last)
    end subroutine validate_values
end program test_array_constructor_shape_inference_18_constructor_assisted_reduction_guard_1
"#,
    );
    assert_eq!(out, vec!["1", "1", "1"]);
}

#[test]
fn array_constructor_shape_inference_19_logical_vector_shape_from_constructor() {
    let out = run_prints(
        r#"
program test_array_constructor_shape_inference_19_logical_vector_shape_from_constructor
    logical, allocatable :: values(:)
    values = (/ .true., .false., .true., .true., .false. /)
    print *, size(values)
    print *, count(values)
    print *, merge(1, 0, values(1))
    print *, merge(1, 0, values(5))
end program test_array_constructor_shape_inference_19_logical_vector_shape_from_constructor
"#,
    );
    assert_eq!(out, vec!["5", "3", "1", "0"]);
}

#[test]
fn array_constructor_shape_inference_20_character_vector_shape_inference() {
    let out = run_prints(
        r#"
program test_array_constructor_shape_inference_20_character_vector_shape_inference
    character(len=4), allocatable :: values(:)
    values = (/ 'ab', 'cde', 'x' /)
    print *, size(values)
    print *, len(values(1))
    print *, len_trim(values(2))
    print *, len_trim(values(3))
end program test_array_constructor_shape_inference_20_character_vector_shape_inference
"#,
    );
    assert_eq!(out, vec!["3", "4", "3", "1"]);
}
