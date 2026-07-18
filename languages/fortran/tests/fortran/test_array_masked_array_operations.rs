use super::helpers::run_prints;

#[test]
fn array_masked_array_operations_basic_where_mask() {
    let out = run_prints(
        r#"
program array_masked_array_operations_basic_where_mask
    integer, allocatable :: values(:)
    integer, allocatable :: result(:)
    values = (/ -1, 0, 1, 2, 3 /)
    result = values
    where (values >= 2)
        result = values * 10
    end where
    print *, result(1)
    print *, result(3)
    print *, result(5)
    print *, sum(result)
end program array_masked_array_operations_basic_where_mask
"#,
    );
    assert_eq!(out, vec!["-1", "1", "30", "34"]);
}

#[test]
fn array_masked_array_operations_where_elsewhere() {
    let out = run_prints(
        r#"
program array_masked_array_operations_where_elsewhere
    integer :: values(6)
    integer :: replaced(6)
    values = (/ 1, -2, 3, -4, 5, 0 /)
    where (values > 0)
        replaced = 100
    elsewhere
        replaced = -100
    end where
    print *, replaced(2)
    print *, replaced(3)
    print *, replaced(6)
    print *, sum(replaced)
end program array_masked_array_operations_where_elsewhere
"#,
    );
    assert_eq!(out, vec!["-100", "100", "-100", "0"]);
}

#[test]
fn array_masked_array_operations_where_with_array_mask() {
    let out = run_prints(
        r#"
program array_masked_array_operations_where_with_array_mask
    integer :: values(5)
    integer :: mask(5)
    integer :: result(5)
    values = (/ 4, 5, 6, 7, 8 /)
    mask = (/ 1, 0, 1, 0, 1 /)
    where (mask == 1)
        result = values + 1
    end where
    print *, sum(result)
    print *, result(1)
    print *, result(2)
    print *, result(5)
end program array_masked_array_operations_where_with_array_mask
"#,
    );
    assert_eq!(out, vec!["25", "5", "0", "9"]);
}

#[test]
fn array_masked_array_operations_where_nested_scalar_transform() {
    let out = run_prints(
        r#"
program array_masked_array_operations_where_nested_scalar_transform
    integer :: values(4)
    integer :: result(4)
    values = (/ 2, 3, 4, 5 /)
    where (mod(values,2) == 0)
        where (values > 3)
            result = values * 3
        elsewhere
            result = values * 2
        end where
    end where
    print *, result(1)
    print *, result(2)
    print *, result(3)
    print *, result(4)
    print *, sum(result)
end program array_masked_array_operations_where_nested_scalar_transform
"#,
    );
    assert_eq!(out, vec!["4", "6", "12", "15", "37"]);
}

#[test]
fn array_masked_array_operations_masked_copy_from_expression() {
    let out = run_prints(
        r#"
program array_masked_array_operations_masked_copy_from_expression
    integer :: values(6)
    integer :: result(6)
    values = (/ 1, 2, 3, 4, 5, 6 /)
    where (values >= 3)
        result = values + 10
    end where
    print *, sum(result)
    print *, result(1)
    print *, result(4)
    print *, result(6)
end program array_masked_array_operations_masked_copy_from_expression
"#,
    );
    assert_eq!(out, vec!["47", "0", "14", "16"]);
}

#[test]
fn array_masked_array_operations_masked_elementary_sum() {
    let out = run_prints(
        r#"
program array_masked_array_operations_masked_elementary_sum
    integer :: values(5)
    integer :: masked_sum
    values = (/ 10, 20, 5, 15, 8 /)
    masked_sum = sum(values, values > 9)
    print *, masked_sum
    print *, count(values > 9)
end program array_masked_array_operations_masked_elementary_sum
"#,
    );
    assert_eq!(out, vec!["45", "4"]);
}

#[test]
fn array_masked_array_operations_merge_rewrite_pattern() {
    let out = run_prints(
        r#"
program array_masked_array_operations_merge_rewrite_pattern
    integer :: values(4)
    integer :: result(4)
    values = (/ 1, 2, 3, 4 /)
    result = merge(values*3, values, values > 2)
    print *, sum(result)
    print *, result(2)
    print *, result(3)
end program array_masked_array_operations_merge_rewrite_pattern
"#,
    );
    assert_eq!(out, vec!["16", "2", "9"]);
}

#[test]
fn array_masked_array_operations_where_mask_with_stride() {
    let out = run_prints(
        r#"
program array_masked_array_operations_where_mask_with_stride
    integer :: values(6)
    integer :: result(6)
    integer :: i
    values = (/ 1, 2, 3, 4, 5, 6 /)
    do i = 1, 6
        if (mod(i, 2) == 0) then
            where (values(i) > 0)
                result(i) = values(i) * 5
            end where
        else
            result(i) = values(i)
        end if
    end do
    print *, sum(result)
    print *, result(1)
    print *, result(2)
    print *, result(6)
end program array_masked_array_operations_where_mask_with_stride
"#,
    );
    assert_eq!(out, vec!["46", "1", "10", "30"]);
}

#[test]
fn array_masked_array_operations_masked_update_inside_do() {
    let out = run_prints(
        r#"
program array_masked_array_operations_masked_update_inside_do
    integer :: values(4)
    integer :: result(4)
    integer :: i
    values = (/ 4, 8, 12, 16 /)
    do i = 1, 4
        if (i <= 2) then
            result(i) = values(i) + 1
        else
            result(i) = values(i) - 1
        end if
    end do
    print *, sum(result)
    print *, result(1)
    print *, result(4)
end program array_masked_array_operations_masked_update_inside_do
"#,
    );
    assert_eq!(out, vec!["39", "5", "15"]);
}

#[test]
fn array_masked_array_operations_mask_for_logical_condition() {
    let out = run_prints(
        r#"
program array_masked_array_operations_mask_for_logical_condition
    integer :: values(7)
    integer :: selected(7)
    integer :: hits
    values = (/ 2, 4, 6, 8, 10, 12, 14 /)
    selected = 0
    where (mod(values, 4) == 0)
        selected = 1
    end where
    hits = sum(selected)
    print *, hits
    print *, count(selected == 1)
    print *, selected(2)
    print *, selected(3)
end program array_masked_array_operations_mask_for_logical_condition
"#,
    );
    assert_eq!(out, vec!["3", "3", "1", "0"]);
}

#[test]
fn array_masked_array_operations_masked_negation_set() {
    let out = run_prints(
        r#"
program array_masked_array_operations_masked_negation_set
    integer :: values(5)
    integer :: result(5)
    values = (/ 3, 7, 11, 13, 17 /)
    where (values >= 10)
        result = 0
    else where
        result = values
    end where
    print *, sum(result)
    print *, result(1)
    print *, result(3)
    print *, result(5)
end program array_masked_array_operations_masked_negation_set
"#,
    );
    assert_eq!(out, vec!["20", "3", "0", "0"]);
}

#[test]
fn array_masked_array_operations_categorical_masked_counts() {
    let out = run_prints(
        r#"
program array_masked_array_operations_categorical_masked_counts
    integer :: values(8)
    integer :: cat_a
    integer :: cat_b
    values = (/ 1, 2, 3, 4, 5, 6, 7, 8 /)
    cat_a = count(values <= 4)
    cat_b = count(values > 4)
    print *, cat_a
    print *, cat_b
    print *, sum(merge(1, 0, values < 3))
end program array_masked_array_operations_categorical_masked_counts
"#,
    );
    assert_eq!(out, vec!["4", "4", "2"]);
}

#[test]
fn array_masked_array_operations_masked_abs_transform() {
    let out = run_prints(
        r#"
program array_masked_array_operations_masked_abs_transform
    integer :: values(6)
    integer :: result(6)
    values = (/ -6, 5, -4, 3, -2, 1 /)
    where (values < 0)
        result = -values
    else where
        result = values
    end where
    print *, result(1)
    print *, result(2)
    print *, result(3)
    print *, sum(result)
end program array_masked_array_operations_masked_abs_transform
"#,
    );
    assert_eq!(out, vec!["6", "5", "4", "21"]);
}

#[test]
fn array_masked_array_operations_masked_reduction_chain() {
    let out = run_prints(
        r#"
program array_masked_array_operations_masked_reduction_chain
    integer :: values(4)
    integer :: result(4)
    values = (/ 5, 10, 15, 20 /)
    where (values >= 10)
        result = values / 5
    end where
    print *, sum(result)
    print *, count(result == 0)
    print *, sum(merge(1, 0, result /= 0))
end program array_masked_array_operations_masked_reduction_chain
"#,
    );
    assert_eq!(out, vec!["9", "2", "3"]);
}

#[test]
fn array_masked_array_operations_where_for_section_copy() {
    let out = run_prints(
        r#"
program array_masked_array_operations_where_for_section_copy
    integer :: source(1:6)
    integer :: result(1:6)
    source = (/ 9, 8, 7, 6, 5, 4 /)
    where (source > 6)
        result = source(1:6)
    elsewhere
        result = 0
    end where
    print *, sum(result)
    print *, result(1)
    print *, result(4)
    print *, result(6)
end program array_masked_array_operations_where_for_section_copy
"#,
    );
    assert_eq!(out, vec!["32", "9", "6", "0"]);
}

#[test]
fn array_masked_array_operations_nested_where_in_construct() {
    let out = run_prints(
        r#"
program array_masked_array_operations_nested_where_in_construct
    integer :: values(5)
    integer :: result(5)
    values = (/ 12, 11, 10, 9, 8 /)
    if (all(values > 0)) then
        where (mod(values, 2) == 0)
            result = 1
        elsewhere
            result = 0
        end where
    else
        result = -1
    end if
    print *, sum(result)
    print *, result(1)
    print *, result(2)
end program array_masked_array_operations_nested_where_in_construct
"#,
    );
    assert_eq!(out, vec!["3", "1", "0"]);
}

#[test]
fn array_masked_array_operations_masked_minmax_mix() {
    let out = run_prints(
        r#"
program array_masked_array_operations_masked_minmax_mix
    integer :: values(6)
    integer :: result(6)
    values = (/ 4, 1, 8, 2, 16, 3 /)
    result = -1
    where (values > 5)
        result = values
    end where
    print *, sum(result)
    print *, result(1)
    print *, result(3)
    print *, result(6)
end program array_masked_array_operations_masked_minmax_mix
"#,
    );
    assert_eq!(out, vec!["20", "-1", "8", "-1"]);
}

#[test]
fn array_masked_array_operations_masked_redundant_mask() {
    let out = run_prints(
        r#"
program array_masked_array_operations_masked_redundant_mask
    integer :: values(4)
    integer :: result(4)
    values = (/ 2, 4, 6, 8 /)
    where (values > 3)
        result = values / 2
    else where
        result = values + 10
    end where
    print *, result(1)
    print *, result(2)
    print *, result(3)
    print *, sum(result)
end program array_masked_array_operations_masked_redundant_mask
"#,
    );
    assert_eq!(out, vec!["12", "2", "3", "20"]);
}

#[test]
fn array_masked_array_operations_masked_scalar_broadcast() {
    let out = run_prints(
        r#"
program array_masked_array_operations_masked_scalar_broadcast
    integer :: values(3)
    integer :: result(3)
    values = (/ 1, 2, 3 /)
    where (values /= 0)
        result = 7
    end where
    print *, sum(result)
    print *, result(1)
    print *, result(3)
end program array_masked_array_operations_masked_scalar_broadcast
"#,
    );
    assert_eq!(out, vec!["21", "7", "7"]);
}
