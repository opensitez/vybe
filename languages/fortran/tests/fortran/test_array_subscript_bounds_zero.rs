use super::helpers::run_prints;

#[test]
fn array_subscript_bounds_zero_basic_bounds() {
    let out = run_prints(
        r#"
program array_subscript_bounds_zero_basic_bounds
    integer :: values(0:4)
    values = (/10, 11, 12, 13, 14/)
    print *, lbound(values)
    print *, ubound(values)
    print *, values(0)
    print *, values(4)
end program array_subscript_bounds_zero_basic_bounds
"#,
    );
    assert_eq!(out, vec!["0", "4", "10", "14"]);
}

#[test]
fn array_subscript_bounds_zero_lowered_indexing() {
    let out = run_prints(
        r#"
program array_subscript_bounds_zero_lowered_indexing
    integer :: values(-2:2)
    integer :: i
    values = (/1, 2, 3, 4, 5/)
    i = values(-2) + values(-1) + values(0)
    print *, i
    print *, values(2)
end program array_subscript_bounds_zero_lowered_indexing
"#,
    );
    assert_eq!(out, vec!["6", "4"]);
}

#[test]
fn array_subscript_bounds_zero_section_start_at_zero() {
    let out = run_prints(
        r#"
program array_subscript_bounds_zero_section_start_at_zero
    integer :: values(0:5)
    values = (/0, 1, 2, 3, 4, 5/)
    print *, sum(values(0:3))
    print *, size(values(0:3))
    print *, values(0:3)(1)
    print *, values(0:3)(4)
end program array_subscript_bounds_zero_section_start_at_zero
"#,
    );
    assert_eq!(out, vec!["3", "4", "0", "3"]);
}

#[test]
fn array_subscript_bounds_zero_bidirectional_sections() {
    let out = run_prints(
        r#"
program array_subscript_bounds_zero_bidirectional_sections
    integer :: values(0:9)
    values = (/1, 2, 3, 4, 5, 6, 7, 8, 9, 10/)
    print *, values(2:8:2)(1)
    print *, values(8:2:-3)(2)
    print *, size(values(2:8:2))
end program array_subscript_bounds_zero_bidirectional_sections
"#,
    );
    assert_eq!(out, vec!["2", "8", "4"]);
}

#[test]
fn array_subscript_bounds_zero_matrix_zero_row() {
    let out = run_prints(
        r#"
program array_subscript_bounds_zero_matrix_zero_row
    integer :: matrix(0:2, 0:2)
    integer :: r
    matrix = reshape((/1, 2, 3, 4, 5, 6, 7, 8, 9/), (/3, 3/))
    r = matrix(0, 0) + matrix(2, 2)
    print *, r
    print *, lbound(matrix, 1)
    print *, lbound(matrix, 2)
end program array_subscript_bounds_zero_matrix_zero_row
"#,
    );
    assert_eq!(out, vec!["10", "0", "0"]);
}

#[test]
fn array_subscript_bounds_zero_char_length_preserved() {
    let out = run_prints(
        r#"
program array_subscript_bounds_zero_char_length_preserved
    character(len=3) :: items(0:3)
    items = (/'aaa', 'bbb', 'ccc', 'ddd'/)
    print *, trim(items(0))
    print *, trim(items(3))
    print *, size(items)
end program array_subscript_bounds_zero_char_length_preserved
"#,
    );
    assert_eq!(out, vec!["aaa", "ddd", "4"]);
}

#[test]
fn array_subscript_bounds_zero_lower_bound_includes_zero_index() {
    let out = run_prints(
        r#"
program array_subscript_bounds_zero_lower_bound_includes_zero_index
    integer :: values(0:1)
    values = (/99, 100/)
    values(0) = values(0) + 1
    print *, values(0)
    print *, values(1)
end program array_subscript_bounds_zero_lower_bound_includes_zero_index
"#,
    );
    assert_eq!(out, vec!["100", "100"]);
}

#[test]
fn array_subscript_bounds_zero_zero_length_not_overflowed() {
    let out = run_prints(
        r#"
program array_subscript_bounds_zero_zero_length_not_overflowed
    integer :: values(0:-1)
    print *, size(values)
end program array_subscript_bounds_zero_zero_length_not_overflowed
"#,
    );
    assert_eq!(out, vec!["0"]);
}
