use super::helpers::run_prints;

#[test]
fn array_implied_do_ordering_linear_fill() {
    let out = run_prints(
        r#"
program array_implied_do_ordering_linear_fill
    integer :: values(5)
    values = [(i, i=1,5)]
    print *, values(1)
    print *, values(5)
    print *, sum(values)
end program array_implied_do_ordering_linear_fill
"#,
    );
    assert_eq!(out, vec!["1", "5", "15"]);
}

#[test]
fn array_implied_do_ordering_descending_fill() {
    let out = run_prints(
        r#"
program array_implied_do_ordering_descending_fill
    integer :: values(4)
    values = [(i, i=8,2,-2)]
    print *, size(values)
    print *, values(1)
    print *, values(4)
    print *, sum(values)
end program array_implied_do_ordering_descending_fill
"#,
    );
    assert_eq!(out, vec!["4", "8", "2", "20"]);
}

#[test]
fn array_implied_do_ordering_stride_fill() {
    let out = run_prints(
        r#"
program array_implied_do_ordering_stride_fill
    integer :: values(4)
    values = [(i, i=1,8,2)]
    print *, values(1)
    print *, values(4)
    print *, sum(values)
end program array_implied_do_ordering_stride_fill
"#,
    );
    assert_eq!(out, vec!["1", "7", "16"]);
}

#[test]
fn array_implied_do_ordering_nested_index_expression() {
    let out = run_prints(
        r#"
program array_implied_do_ordering_nested_index_expression
    integer :: values(5)
    values = [(i*2, i = 1,5)]
    print *, values(1)
    print *, values(5)
    print *, sum(values)
end program array_implied_do_ordering_nested_index_expression
"#,
    );
    assert_eq!(out, vec!["2", "10", "30"]);
}

#[test]
fn array_implied_do_ordering_expression_with_offset() {
    let out = run_prints(
        r#"
program array_implied_do_ordering_expression_with_offset
    integer :: values(5)
    values = [(i + 3, i = 0, 4)]
    print *, values(1)
    print *, values(3)
    print *, values(5)
    print *, sum(values)
end program array_implied_do_ordering_expression_with_offset
"#,
    );
    assert_eq!(out, vec!["3", "5", "7", "35"]);
}

#[test]
fn array_implied_do_ordering_nested_construct_stable_order_1() {
    let out = run_prints(
        r#"
program array_implied_do_ordering_nested_construct_stable_order_1
    integer :: values(6)
    integer :: i
    i = size([(j, j = 1, 6)])
    values = [(i, i = 1, 6)]
    print *, i
    print *, values(1)
    print *, values(6)
    print *, sum(values)
end program array_implied_do_ordering_nested_construct_stable_order_1
"#,
    );
    assert_eq!(out, vec!["6", "1", "6", "21"]);
}

#[test]
fn array_implied_do_ordering_implied_do_of_logical_mask() {
    let out = run_prints(
        r#"
program array_implied_do_ordering_implied_do_of_logical_mask
    integer :: values(4)
    values = [(merge(1, 0, mod(i,2) == 0), i = 1, 4)]
    print *, sum(values)
    print *, values(2)
    print *, values(4)
end program array_implied_do_ordering_implied_do_of_logical_mask
"#,
    );
    assert_eq!(out, vec!["2", "1", "1"]);
}

#[test]
fn array_implied_do_ordering_nested_roundtrip() {
    let out = run_prints(
        r#"
program array_implied_do_ordering_nested_roundtrip
    integer :: values(6)
    values = [(i+j, i=1,2, j=1,3)]
    print *, size(values)
    print *, sum(values)
    print *, values(1)
    print *, values(6)
end program array_implied_do_ordering_nested_roundtrip
"#,
    );
    assert_eq!(out, vec!["6", "21", "2", "5"]);
}

#[test]
fn array_implied_do_ordering_nested_order_swap() {
    let out = run_prints(
        r#"
program array_implied_do_ordering_nested_order_swap
    integer :: values(6)
    values = [(i+j, j=1,3, i=1,2)]
    print *, sum(values)
    print *, values(1)
    print *, values(6)
end program array_implied_do_ordering_nested_order_swap
"#,
    );
    assert_eq!(out, vec!["21", "2", "5"]);
}

#[test]
fn array_implied_do_ordering_zero_stride_guarded() {
    let out = run_prints(
        r#"
program array_implied_do_ordering_zero_stride_guarded
    integer :: status
    integer :: values(3)
    if (1 <= 3) then
        values = [(1, i = 1, 3)]
    else
        values = 0
    end if
    status = sum(values)
    print *, status
    print *, values(3)
end program array_implied_do_ordering_zero_stride_guarded
"#,
    );
    assert_eq!(out, vec!["3", "1"]);
}

#[test]
fn array_implied_do_ordering_constructed_from_runtime_bound() {
    let out = run_prints(
        r#"
program array_implied_do_ordering_constructed_from_runtime_bound
    integer :: n
    integer :: values(4)
    n = 1
    values = [(i*n, i=1,4)]
    print *, sum(values)
    print *, values(4)
    print *, values(2)
end program array_implied_do_ordering_constructed_from_runtime_bound
"#,
    );
    assert_eq!(out, vec!["10", "4", "2"]);
}

#[test]
fn array_implied_do_ordering_sectioned_implied_do_fill() {
    let out = run_prints(
        r#"
program array_implied_do_ordering_sectioned_implied_do_fill
    integer :: a(1:8)
    a = [(i, i = 1, 8)]
    a(3:6) = [(i*2, i = 1,4)]
    print *, a(1)
    print *, a(5)
    print *, sum(a)
end program array_implied_do_ordering_sectioned_implied_do_fill
"#,
    );
    assert_eq!(out, vec!["1", "8", "34"]);
}

#[test]
fn array_implied_do_ordering_multi_expression_series() {
    let out = run_prints(
        r#"
program array_implied_do_ordering_multi_expression_series
    integer :: values(3)
    values = [(i*i + 1, i = 1, 3)]
    print *, values(1)
    print *, values(2)
    print *, values(3)
    print *, sum(values)
end program array_implied_do_ordering_multi_expression_series
"#,
    );
    assert_eq!(out, vec!["2", "5", "10", "17"]);
}

#[test]
fn array_implied_do_ordering_reverse_then_forward() {
    let out = run_prints(
        r#"
program array_implied_do_ordering_reverse_then_forward
    integer :: values(4)
    values = [(i, i=4,1,-1)]
    print *, values(1)
    print *, values(4)
    print *, sum(values)
end program array_implied_do_ordering_reverse_then_forward
"#,
    );
    assert_eq!(out, vec!["4", "1", "10"]);
}

#[test]
fn array_implied_do_ordering_implied_fill_for_real() {
    let out = run_prints(
        r#"
program array_implied_do_ordering_implied_fill_for_real
    real, allocatable :: values(:)
    integer :: n
    values = [(1.0 * i / 2.0, i = 1,4)]
    n = nint(sum(values) * 10.0)
    print *, n
    print *, nint(values(1)*10)
    print *, nint(values(4)*10)
end program array_implied_do_ordering_implied_fill_for_real
"#,
    );
    assert_eq!(out, vec!["50", "5", "20"]);
}

#[test]
fn array_implied_do_ordering_complex_sequence() {
    let out = run_prints(
        r#"
program array_implied_do_ordering_complex_sequence
    integer :: values(5)
    values = [(i*3 + i/2, i = 1, 5)]
    print *, values(1)
    print *, values(5)
    print *, sum(values)
end program array_implied_do_ordering_complex_sequence
"#,
    );
    assert_eq!(out, vec!["4", "17", "50"]);
}

#[test]
fn array_implied_do_ordering_rebind_between_fills() {
    let out = run_prints(
        r#"
program array_implied_do_ordering_rebind_between_fills
    integer :: a(4)
    integer :: b(4)
    a = [(i, i = 1,4)]
    b = [(i*2, i = 1,4)]
    print *, sum(a)
    print *, sum(b)
    print *, b(2) - a(2)
end program array_implied_do_ordering_rebind_between_fills
"#,
    );
    assert_eq!(out, vec!["10", "20", "2"]);
}

#[test]
fn array_implied_do_ordering_fill_via_function_like_expression() {
    let out = run_prints(
        r#"
program array_implied_do_ordering_fill_via_function_like_expression
    integer :: values(4)
    values = [(i + 10, i = 1,4)]
    print *, sum(values)
    print *, values(1)
    print *, values(4)
end program array_implied_do_ordering_fill_via_function_like_expression
"#,
    );
    assert_eq!(out, vec!["54", "11", "14"]);
}

#[test]
fn array_implied_do_ordering_sectioned_nested_fill_guarded() {
    let out = run_prints(
        r#"
program array_implied_do_ordering_sectioned_nested_fill_guarded
    integer :: values(6)
    integer :: n
    n = 3
    values = 0
    if (n == 3) values = [(i*i, i = 1,6)]
    print *, sum(values)
    print *, values(3)
    print *, values(6)
end program array_implied_do_ordering_sectioned_nested_fill_guarded
"#,
    );
    assert_eq!(out, vec!["91", "9", "36"]);
}

#[test]
fn array_implied_do_ordering_stride_then_masked_fill() {
    let out = run_prints(
        r#"
program array_implied_do_ordering_stride_then_masked_fill
    integer :: values(4)
    values = [(i, i = 2,8,2)]
    values = merge(values, 0, values > 2)
    print *, sum(values)
    print *, values(1)
    print *, values(2)
    print *, values(3)
    print *, values(4)
end program array_implied_do_ordering_stride_then_masked_fill
"#,
    );
    assert_eq!(out, vec!["6", "2", "0", "0", "0"]);
}

