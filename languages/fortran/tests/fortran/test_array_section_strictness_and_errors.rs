use super::helpers::{parse_ok, run_prints};

#[test]
fn array_section_strictness_and_errors_empty_forward_triplet_has_zero_size() {
    let out = run_prints(
        r#"
program array_section_strictness_and_errors_empty_forward_triplet_has_zero_size
    integer :: values(1:5)
    values = (/1, 2, 3, 4, 5/)
    print *, size(values(5:4))
end program array_section_strictness_and_errors_empty_forward_triplet_has_zero_size
"#,
    );

    assert_eq!(out, vec!["0"]);
}

#[test]
fn array_section_strictness_and_errors_empty_reverse_triplet_has_zero_size() {
    let out = run_prints(
        r#"
program array_section_strictness_and_errors_empty_reverse_triplet_has_zero_size
    integer :: values(1:5)
    values = (/1, 2, 3, 4, 5/)
    print *, size(values(1:5:-1))
end program array_section_strictness_and_errors_empty_reverse_triplet_has_zero_size
"#,
    );

    assert_eq!(out, vec!["0"]);
}

#[test]
fn array_section_strictness_and_errors_lbound_omitted() {
    let out = run_prints(
        r#"
program array_section_strictness_and_errors_lbound_omitted
    integer :: values(0:4)
    values = (/1, 2, 3, 4, 5/)
    print *, lbound(values(:3))
    print *, ubound(values(:3))
    print *, size(values(:3))
end program array_section_strictness_and_errors_lbound_omitted
"#,
    );

    assert_eq!(out, vec!["0", "2", "3"]);
}

#[test]
fn array_section_strictness_and_errors_ubound_omitted() {
    let out = run_prints(
        r#"
program array_section_strictness_and_errors_ubound_omitted
    integer :: values(1:6)
    values = (/1, 2, 3, 4, 5, 6/)
    print *, lbound(values(3:))
    print *, ubound(values(3:))
    print *, size(values(3:))
end program array_section_strictness_and_errors_ubound_omitted
"#,
    );

    assert_eq!(out, vec!["3", "6", "4"]);
}

#[test]
fn array_section_strictness_and_errors_negative_stride_subsection_shape() {
    let out = run_prints(
        r#"
program array_section_strictness_and_errors_negative_stride_subsection_shape
    integer :: values(1:9)
    values = (/1, 2, 3, 4, 5, 6, 7, 8, 9/)
    print *, size(values(9:1:-2))
    print *, values(9:1:-2)(1)
    print *, values(9:1:-2)(3)
    print *, sum(values(9:1:-2))
end program array_section_strictness_and_errors_negative_stride_subsection_shape
"#,
    );

    assert_eq!(out, vec!["5", "9", "5", "25"]);
}

#[test]
fn array_section_strictness_and_errors_matrix_column_stride_bounds() {
    let out = run_prints(
        r#"
program array_section_strictness_and_errors_matrix_column_stride_bounds
    integer :: values(4, 4)
    values = reshape((/ (i, i = 1, 16) /), (/4, 4/))
    print *, size(values(:, 2:4:2), 1)
    print *, size(values(:, 2:4:2), 2)
    print *, sum(values(:, 2:4:2))
end program array_section_strictness_and_errors_matrix_column_stride_bounds
"#,
    );

    assert_eq!(out, vec!["4", "2", "30"]);
}

#[test]
fn array_section_strictness_and_errors_reordered_bounds_preserve_extent() {
    let out = run_prints(
        r#"
program array_section_strictness_and_errors_reordered_bounds_preserve_extent
    integer :: values(1:6)
    values = (/1, 2, 3, 4, 5, 6/)
    print *, values(6:2:-2)(1)
    print *, values(6:2:-2)(2)
    print *, size(values(6:2:-2))
end program array_section_strictness_and_errors_reordered_bounds_preserve_extent
"#,
    );

    assert_eq!(out, vec!["6", "2", "3"]);
}

#[test]
fn array_section_strictness_and_errors_zero_stride_is_rejected_by_parser() {
    assert!(!parse_ok(
        "program array_section_strictness_and_errors_zero_stride_is_rejected_by_parser\n\
            integer :: values(1:5)\n\
            print *, values(1:5:0)\n\
        end program\n",
    ));
}

#[test]
fn array_section_strictness_and_errors_nonmonotonic_upper_bound_is_rejected_by_parser() {
    assert!(!parse_ok(
        "program array_section_strictness_and_errors_nonmonotonic_upper_bound_is_rejected_by_parser\n\
            integer :: values(1:5)\n\
            print *, values(1:3:2:1)\n\
        end program\n",
    ));
}
