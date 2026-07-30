//! Implied-DO coverage: array constructors, DATA implied-do, print implied-do,
//! strides, and nested forms. Distinct from `test_legacy.rs` (basic DATA implied-do),
//! `test_legacy_data_extended.rs` (DATA compile-only extensions), and
//! `test_arrays.rs` (single compile-only array constructor).

use super::helpers::compile_ok;

fortran_cases! {
    // ── Array constructors (i, i=1,n) ────────────────────────────────

    ac_squares_sum_one_to_four => {
        "program t\ninteger :: a(4) = [(i * i, i = 1, 4)]\nprint *, sum(a)\nend program t\n",
        ["30"]
    };

    ac_old_syntax_slash_paren_sum => {
        "program t\ninteger :: a(4) = (/ (i, i = 1, 4) /)\nprint *, sum(a)\nend program t\n",
        ["10"]
    };

    ac_odd_values_from_linear_expr => {
        "program t\ninteger :: a(4) = [(2 * i + 1, i = 1, 4)]\nprint *, sum(a)\nend program t\n",
        ["24"]
    };

    ac_real_values_from_index => {
        "program t\nreal :: r(3) = [(real(i) + 0.5, i = 1, 3)]\nprint *, r(2)\nend program t\n",
        ["2.5"]
    };

    // ── Strides ────────────────────────────────────────────────────────

    ac_stride_two_sum_one_to_nine => {
        "program t\ninteger :: a(5) = [(i, i = 1, 9, 2)]\nprint *, sum(a)\nend program t\n",
        ["25"]
    };

    ac_descending_stride_sum => {
        "program t\ninteger :: a(5) = [(i, i = 5, 1, -1)]\nprint *, sum(a)\nend program t\n",
        ["15"]
    };

    ac_stride_three_corners => {
        "program t\ninteger :: a(4) = [(i, i = 2, 11, 3)]\nprint *, a(1)\nprint *, a(4)\nend program t\n",
        ["2", "11"]
    };

    ac_zero_origin_stride_four => {
        "program t\ninteger :: a(3) = [(i, i = 0, 8, 4)]\nprint *, sum(a)\nend program t\n",
        ["12"]
    };

    ac_zero_to_five_even_stride => {
        "program t\ninteger :: a(3) = [(i, i = 0, 5, 2)]\nprint *, sum(a)\nend program t\n",
        ["6"]
    };

    ac_descending_implied_do_bound => {
        "program t\ninteger, parameter :: n = 6\ninteger :: a(4) = [(i, i = n, 1, -2)]\nprint *, a(1)\nprint *, a(4)\nend program t\n",
        ["6", "2"]
    };

    ac_nested_implied_do_sums_rows => {
        "program t\ninteger :: a(6) = [((i*10 + j, j = 1, 2), i = 1, 3)]\nprint *, sum(a)\nend program t\n",
        ["129"]
    };

    ac_nested_implied_do_with_offset_and_stride => {
        "program t\ninteger :: a(4) = [((i + j, j = 0, 2, 2), i = 1, 2)]\nprint *, sum(a)\nend program t\n",
        ["10"]
    };

    ac_implied_do_with_parameter_bound => {
        "program t\ninteger, parameter :: n = 4\ninteger :: a(n) = [(i, i = 1, n)]\nprint *, sum(a)\nend program t\n",
        ["10"]
    };

    // ── Print with implied-do array constructors ───────────────────────

    print_list_directed_from_implied_do => {
        "program t\nprint *, [(i, i = 1, 5)]\nend program t\n",
        ["1,2,3,4,5"]
    };

    print_old_syntax_implied_do_stride => {
        "program t\nprint *, (/ (i, i = 2, 8, 3) /)\nend program t\n",
        ["2,5,8"]
    };

    print_squared_implied_do_values => {
        "program t\nprint *, [(i * i, i = 1, 4)]\nend program t\n",
        ["1,4,9,16"]
    };
}

// ── DATA implied-do (compile-only; distinct from legacy suites) ──────

#[test]
fn data_implied_do_second_matrix_row() {
    compile_ok(
        r#"
program t
    integer :: grid(2, 3)
    data (grid(2, j), j = 1, 3) /10, 20, 30/
    print *, grid(2, 2)
end program t
"#,
    );
}

#[test]
fn data_implied_do_zero_origin_stride_four() {
    compile_ok(
        r#"
program t
    integer :: slots(3)
    data (slots(i), i = 0, 8, 4) /0, 4, 8/
    print *, slots(2)
end program t
"#,
    );
}

#[test]
fn data_implied_do_mid_range_stride_four() {
    compile_ok(
        r#"
program t
    integer :: buf(3)
    data (buf(i), i = 3, 11, 4) /11, 15, 19/
    print *, buf(3)
end program t
"#,
    );
}

#[test]
fn data_implied_do_nested_indices_two_dimensions() {
    compile_ok(
        r#"
program t
    integer :: mat(2, 3)
    data ((mat(i, j), j = 1, 3), i = 1, 2) /1, 2, 3, 4, 5, 6/
    print *, mat(2, 3)
end program t
"#,
    );
}

#[test]
fn data_implied_do_negative_stride_two_dimensional() {
    compile_ok(
        r#"
program t
    integer :: arr(3)
    data (arr(i), i = 5, 1, -2) /5, 3, 1/
    print *, arr(1)
end program t
"#,
    );
}

#[test]
fn data_implied_do_nested_indices_three_way() {
    compile_ok(
        r#"
program t
    integer :: vol(2,2,2)
    data (((vol(i, j, k), k = 1, 2), j = 1, 2), i = 1, 2) /1,2,3,4,5,6,7,8/
    print *, vol(2, 2, 2)
end program t
"#,
    );
}

#[test]
fn data_implied_do_with_expression_stride() {
    compile_ok(
        r#"
program t
    integer :: v(4)
    data (v(i), i = 2, 8, 2 + 0) /2, 4, 6, 8/
    print *, v(4)
end program t
"#,
    );
}

#[test]
fn data_implied_do_string_length_literal_sequence() {
    compile_ok(
        r#"
program t
    character(len=1) :: tags(3)
    data (tags(i), i = 1, 3) /'a','b','c'/
    print *, tags(2)
end program t
"#,
    );
}

#[test]
fn data_implied_do_with_parameterized_bounds() {
    compile_ok(
        r#"
program t
    integer, parameter :: n = 3
    integer :: out(n)
    data (out(i), i = 1, n) /10,20,30/
    print *, out(2)
end program t
"#,
    );
}
