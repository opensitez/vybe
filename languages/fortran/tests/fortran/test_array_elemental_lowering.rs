use super::helpers::run_prints;

#[test]
fn array_elemental_lowering_abs_on_signed_integers() {
    let out = run_prints(
        r#"
program array_elemental_lowering_abs_on_signed_integers
    integer, allocatable :: values(:)
    values = (/ -5, -1, 0, 2, -3 /)
    print *, sum(abs(values))
    print *, abs(values(1))
    print *, abs(values(5))
end program array_elemental_lowering_abs_on_signed_integers
"#,
    );
    assert_eq!(out, vec!["11", "5", "3"]);
}

#[test]
fn array_elemental_lowering_sign_on_zero_positive_negative() {
    let out = run_prints(
        r#"
program array_elemental_lowering_sign_on_zero_positive_negative
    integer, allocatable :: values(:)
    integer :: positive
    integer :: negative
    values = (/ -4, 0, 7 /)
    positive = sum(sign(1, values))
    negative = count(values < 0)
    print *, positive
    print *, negative
    print *, sign(-1, values(3))
end program array_elemental_lowering_sign_on_zero_positive_negative
"#,
    );
    assert_eq!(out, vec!["1", "1", "-1"]);
}

#[test]
fn array_elemental_lowering_implicit_type_coercion_through_real() {
    let out = run_prints(
        r#"
program array_elemental_lowering_implicit_type_coercion_through_real
    integer, allocatable :: ints(:)
    real, allocatable :: vals(:)
    ints = (/ 1, 2, 3, 4, 5 /)
    vals = real(ints) + 0.5
    print *, sum(ints)
    print *, nint(sum(vals))
    print *, nint(vals(1))
    print *, nint(vals(size(vals)))
end program array_elemental_lowering_implicit_type_coercion_through_real
"#,
    );
    assert_eq!(out, vec!["15", "17", "2", "6"]);
}

#[test]
fn array_elemental_lowering_addition_vectorized() {
    let out = run_prints(
        r#"
program array_elemental_lowering_addition_vectorized
    integer, allocatable :: left(:), right(:)
    left = (/ 1, 2, 3, 4 /)
    right = (/ 4, 3, 2, 1 /)
    print *, sum(left + right)
    print *, (left + right)(1)
    print *, (left + right)(4)
end program array_elemental_lowering_addition_vectorized
"#,
    );
    assert_eq!(out, vec!["20", "5", "5"]);
}

#[test]
fn array_elemental_lowering_multiplication_vectorized() {
    let out = run_prints(
        r#"
program array_elemental_lowering_multiplication_vectorized
    integer, allocatable :: left(:), right(:)
    left = (/ 2, 3, 4, 5 /)
    right = (/ 1, 2, 3, 4 /)
    print *, sum(left * right)
    print *, (left * right)(1)
    print *, (left * right)(3)
end program array_elemental_lowering_multiplication_vectorized
"#,
    );
    assert_eq!(out, vec!["40", "2", "12"]);
}

#[test]
fn array_elemental_lowering_elemental_division_with_constant() {
    let out = run_prints(
        r#"
program array_elemental_lowering_elemental_division_with_constant
    integer, allocatable :: values(:)
    integer, allocatable :: half(:)
    values = (/ 2, 4, 6, 8 /)
    half = values / 2
    print *, size(half)
    print *, sum(half)
    print *, half(4)
end program array_elemental_lowering_elemental_division_with_constant
"#,
    );
    assert_eq!(out, vec!["4", "10", "4"]);
}

#[test]
fn array_elemental_lowering_maxval_on_expression() {
    let out = run_prints(
        r#"
program array_elemental_lowering_maxval_on_expression
    integer, allocatable :: values(:)
    values = (/ -9, 12, 4, 18, 3 /)
    print *, maxval(values)
    print *, maxval(abs(values))
    print *, size(pack(values, values > 5))
end program array_elemental_lowering_maxval_on_expression
"#,
    );
    assert_eq!(out, vec!["18", "18", "2"]);
}

#[test]
fn array_elemental_lowering_minval_and_index() {
    let out = run_prints(
        r#"
program array_elemental_lowering_minval_and_index
    integer, allocatable :: values(:)
    values = (/ 7, 3, 9, 3, 1 /)
    print *, minval(values)
    print *, minval(values) + maxval(values)
    print *, maxloc(values)
end program array_elemental_lowering_minval_and_index
"#,
    );
    assert_eq!(out, vec!["1", "10", "5"]);
}

#[test]
fn array_elemental_lowering_where_mask_assignment() {
    let out = run_prints(
        r#"
program array_elemental_lowering_where_mask_assignment
    integer, allocatable :: values(:)
    integer, allocatable :: marked(:)
    values = (/ 1, 2, 3, 4, 5 /)
    marked = values
    where (values > 3)
        marked = 99
    end where
    print *, sum(marked)
    print *, marked(3)
    print *, marked(5)
end program array_elemental_lowering_where_mask_assignment
"#,
    );
    assert_eq!(out, vec!["209", "3", "99"]);
}

#[test]
fn array_elemental_lowering_merge_based_branching() {
    let out = run_prints(
        r#"
program array_elemental_lowering_merge_based_branching
    integer, allocatable :: values(:)
    integer, allocatable :: merged(:)
    values = (/ -1, 2, -3, 4 /)
    merged = merge(10, 0, values > 0)
    print *, sum(merged)
    print *, merged(1)
    print *, merged(4)
end program array_elemental_lowering_merge_based_branching
"#,
    );
    assert_eq!(out, vec!["20", "0", "10"]);
}

#[test]
fn array_elemental_lowering_pack_unpacked_counts() {
    let out = run_prints(
        r#"
program array_elemental_lowering_pack_unpacked_counts
    integer, allocatable :: values(:)
    integer :: packed_count
    integer :: unpack_count
    values = (/ 5, 0, -2, 8, 0, 1 /)
    packed_count = size(pack(values, values > 0))
    unpack_count = size(pack(values, values == 0))
    print *, packed_count
    print *, unpack_count
    print *, values(1) + values(6)
end program array_elemental_lowering_pack_unpacked_counts
"#,
    );
    assert_eq!(out, vec!["3", "2", "6"]);
}

#[test]
fn array_elemental_lowering_all_and_any_maskes() {
    let out = run_prints(
        r#"
program array_elemental_lowering_all_and_any_maskes
    integer, allocatable :: values(:)
    values = (/ 1, 1, 1, 0 /)
    print *, merge(1, 0, all(values == 1))
    print *, merge(1, 0, any(values == 0))
    print *, count(values /= 1)
end program array_elemental_lowering_all_and_any_maskes
"#,
    );
    assert_eq!(out, vec!["0", "1", "1"]);
}

#[test]
fn array_elemental_lowering_count_with_nested_mask() {
    let out = run_prints(
        r#"
program array_elemental_lowering_count_with_nested_mask
    integer, allocatable :: values(:)
    values = (/ 0, 1, 2, 3, -1, -2 /)
    print *, count(values >= 0)
    print *, count((values == 0) .or. (values == -1))
    print *, maxval(abs(values))
end program array_elemental_lowering_count_with_nested_mask
"#,
    );
    assert_eq!(out, vec!["4", "2", "3"]);
}

#[test]
fn array_elemental_lowering_real_vectorized_math() {
    let out = run_prints(
        r#"
program array_elemental_lowering_real_vectorized_math
    real, allocatable :: values(:)
    integer :: total
    values = (/ 0.5, 1.0, 1.5, 2.0 /)
    total = nint(sum(sin(values) + cos(values)))
    print *, total
    print *, nint(sum(values * 2.0))
end program array_elemental_lowering_real_vectorized_math
"#,
    );
    assert_eq!(out, vec!["3", "8"]);
}

#[test]
fn array_elemental_lowering_logical_from_numeric_comparison() {
    let out = run_prints(
        r#"
program array_elemental_lowering_logical_from_numeric_comparison
    integer, allocatable :: values(:)
    integer :: positives
    values = (/ -3, -1, 0, 4, 2 /)
    positives = count(values > 0)
    print *, positives
    print *, maxval(merge(1, 0, values > 0))
end program array_elemental_lowering_logical_from_numeric_comparison
"#,
    );
    assert_eq!(out, vec!["2", "1"]);
}

#[test]
fn array_elemental_lowering_bit_mask_style_boolean_transform() {
    let out = run_prints(
        r#"
program array_elemental_lowering_bit_mask_style_boolean_transform
    integer, allocatable :: values(:)
    integer, allocatable :: flags(:)
    values = (/ 1, 2, 4, 8 /)
    flags = iand(values, 2)
    print *, sum(flags)
    print *, flags(1)
    print *, flags(4)
    print *, count(flags > 0)
end program array_elemental_lowering_bit_mask_style_boolean_transform
"#,
    );
    assert_eq!(out, vec!["10", "0", "0", "1"]);
}

#[test]
fn array_elemental_lowering_reduction_after_cast() {
    let out = run_prints(
        r#"
program array_elemental_lowering_reduction_after_cast
    real, allocatable :: values(:)
    integer :: count_hi
    values = (/ 1.2, 2.8, 3.6, 4.4 /)
    count_hi = count(int(values) >= 3)
    print *, count_hi
    print *, nint(maxval(values))
    print *, nint(minval(values))
end program array_elemental_lowering_reduction_after_cast
"#,
    );
    assert_eq!(out, vec!["2", "4", "1"]);
}

#[test]
fn array_elemental_lowering_elemental_shift_equivalence() {
    let out = run_prints(
        r#"
program array_elemental_lowering_elemental_shift_equivalence
    integer, allocatable :: values(:)
    integer, allocatable :: copied(:)
    values = (/ 2, 4, 6, 8 /)
    copied = 2 * values
    print *, sum(copied)
    print *, copied(2) / values(1)
    print *, copied(4) - values(4)
end program array_elemental_lowering_elemental_shift_equivalence
"#,
    );
    assert_eq!(out, vec!["40", "4", "8"]);
}

#[test]
fn array_elemental_lowering_sectionwise_copy_like_transform() {
    let out = run_prints(
        r#"
program array_elemental_lowering_sectionwise_copy_like_transform
    integer :: source(1:6)
    integer :: target(1:6)
    source = (/ 1, 2, 3, 4, 5, 6 /)
    target = 0
    target(2:5) = source(2:5)
    print *, source(1)
    print *, target(1)
    print *, sum(target)
    print *, target(2) + target(5)
end program array_elemental_lowering_sectionwise_copy_like_transform
"#,
    );
    assert_eq!(out, vec!["1", "0", "14", "7"]);
}

#[test]
fn array_elemental_lowering_cascade_reassignments() {
    let out = run_prints(
        r#"
program array_elemental_lowering_cascade_reassignments
    integer, allocatable :: values(:)
    integer :: result
    values = (/ 1, 2, 3, 4 /)
    values = values + 1
    values = values * 2
    result = sum(values)
    print *, result
    print *, values(1)
    print *, values(4)
end program array_elemental_lowering_cascade_reassignments
"#,
    );
    assert_eq!(out, vec!["20", "4", "10"]);
}

#[test]
fn array_elemental_lowering_elemental_reshape_replay() {
    let out = run_prints(
        r#"
program array_elemental_lowering_elemental_reshape_replay
    integer, allocatable :: flat(:)
    integer, allocatable :: matrix(:,:)
    integer :: corner
    flat = (/ 1, 2, 3, 4, 5, 6 /)
    matrix = reshape(abs(flat), (/2, 3/))
    corner = matrix(1,1) + matrix(2,3)
    print *, size(matrix,1)
    print *, size(matrix,2)
    print *, corner
    print *, sum(matrix)
end program array_elemental_lowering_elemental_reshape_replay
"#,
    );
    assert_eq!(out, vec!["2", "3", "7", "21"]);
}

#[test]
fn array_elemental_lowering_modular_remainder_vector() {
    let out = run_prints(
        r#"
program array_elemental_lowering_modular_remainder_vector
    integer, allocatable :: values(:)
    integer, allocatable :: rems(:)
    values = (/ 8, 9, 10, 11, 12 /)
    rems = mod(values, 5)
    print *, size(rems)
    print *, sum(rems)
    print *, rems(2)
    print *, rems(5)
end program array_elemental_lowering_modular_remainder_vector
"#,
    );
    assert_eq!(out, vec!["5", "10", "4", "2"]);
}

#[test]
fn array_elemental_lowering_overlapping_section_with_rhs_temp_semantics() {
    let out = run_prints(
        r#"
program array_elemental_lowering_overlapping_section_with_rhs_temp_semantics
    integer, allocatable :: values(:)
    values = (/ 1, 2, 3, 4, 5, 6 /)
    values(2:5) = values(1:4) + values(2:5)
    print *, size(values)
    print *, sum(values)
    print *, values(1)
    print *, values(2)
    print *, values(5)
end program array_elemental_lowering_overlapping_section_with_rhs_temp_semantics
"#,
    );
    assert_eq!(out, vec!["6", "31", "1", "3", "9"]);
}

#[test]
fn array_elemental_lowering_real_where_numeric_to_logical() {
    let out = run_prints(
        r#"
program array_elemental_lowering_real_where_numeric_to_logical
    real, allocatable :: values(:)
    logical, allocatable :: flags(:)
    values = (/ -1.0, 0.5, 2.25, -0.2 /)
    flags = values > 0.0
    print *, size(flags)
    print *, count(flags)
    print *, merge(1, 0, all(flags(1:2)))
    print *, merge(1, 0, any(flags))
    print *, merge(1, 0, flags(3))
end program array_elemental_lowering_real_where_numeric_to_logical
"#,
    );
    assert_eq!(out, vec!["4", "2", "0", "1", "1"]);
}
