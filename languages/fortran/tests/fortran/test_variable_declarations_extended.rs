//! Extended Fortran variable declarations: typed kinds, dimension attributes,
//! parameter statements, initialization, implicit none, double precision,
//! and complex declarations.

fortran_cases! {
    // ── Integer with kind ────────────────────────────────────────────

    integer_kind_2_init => {
        "program t\nimplicit none\ninteger(kind=2) :: s = 300\nprint *, s\nend program t\n",
        ["300"]
    };

    integer_kind_4_init => {
        "program t\nimplicit none\ninteger(kind=4) :: x = 17\nprint *, x\nend program t\n",
        ["17"]
    };

    integer_kind_8_suffix => {
        "program t\nimplicit none\ninteger(kind=8) :: big = 42_8\nprint *, big\nend program t\n",
        ["42"]
    };

    integer_kind_from_parameter => {
        "program t\nimplicit none\ninteger, parameter :: ik = 4\ninteger(kind=ik) :: n = 99\nprint *, n\nend program t\n",
        ["99"]
    };

    integer_selected_int_kind_decl => {
        "program t\nimplicit none\ninteger, parameter :: ik = selected_int_kind(9)\ninteger(kind=ik) :: v = 512\nprint *, v\nend program t\n",
        ["512"]
    };

    // ── Real with kind ─────────────────────────────────────────────────

    real_kind_4_init => {
        "program t\nimplicit none\nreal(kind=4) :: x = 2.5\nprint *, x\nend program t\n",
        ["2.5"]
    };

    real_kind_8_init => {
        "program t\nimplicit none\nreal(kind=8) :: d = 1.25_8\nprint *, d\nend program t\n",
        ["1.25"]
    };

    real_kind_from_parameter => {
        "program t\nimplicit none\ninteger, parameter :: rk = 8\nreal(kind=rk) :: y = 3.5_8\nprint *, y\nend program t\n",
        ["3.5"]
    };

    real_selected_real_kind_decl => {
        "program t\nimplicit none\ninteger, parameter :: rk = selected_real_kind(6)\nreal(kind=rk) :: z = 0.125\nprint *, z\nend program t\n",
        ["0.125"]
    };

    // ── Logical with kind ──────────────────────────────────────────────

    logical_kind_true_init => {
        "program t\nimplicit none\nlogical(kind=4) :: flag = .true.\nprint *, flag\nend program t\n",
        ["true"]
    };

    logical_kind_false_init => {
        "program t\nimplicit none\nlogical(kind=4) :: flag = .false.\nprint *, flag\nend program t\n",
        ["false"]
    };

    logical_kind_from_not_expression => {
        "program t\nimplicit none\nlogical :: a = .true.\nlogical :: b = .not. a\nprint *, b\nend program t\n",
        ["false"]
    };

    logical_parameter_constant => {
        "program t\nimplicit none\nlogical, parameter :: ok = .true.\nprint *, ok\nend program t\n",
        ["true"]
    };

    // ── Character with kind and length ─────────────────────────────────

    character_len_1_init => {
        "program t\nimplicit none\ncharacter(len=1) :: c = \"Z\"\nprint *, c\nend program t\n",
        ["Z"]
    };

    character_len_8_trim_print => {
        "program t\nimplicit none\ncharacter(len=8) :: s = \"fortran\"\nprint *, trim(s)\nend program t\n",
        ["fortran"]
    };

    character_kind_1_len_3 => {
        "program t\nimplicit none\ncharacter(kind=1, len=3) :: tag = \"abc\"\nprint *, tag\nend program t\n",
        ["abc"]
    };

    character_parameter_const => {
        "program t\nimplicit none\ncharacter(len=5), parameter :: greeting = \"hello\"\nprint *, greeting\nend program t\n",
        ["hello"]
    };

    character_array_decl => {
        "program t\nimplicit none\ncharacter(len=4) :: names(2)\nnames(1) = \"alpha\"\nnames(2) = \"beta\"\nprint *, trim(names(2))\nend program t\n",
        ["beta"]
    };

    // ── Dimension attributes ───────────────────────────────────────────

    dimension_1d_integer_fill => {
        "program t\nimplicit none\ninteger, dimension(4) :: a\ninteger :: i\ndo i = 1, 4\n  a(i) = i * 10\nend do\nprint *, a(3)\nend program t\n",
        ["30"]
    };

    dimension_2d_corner_element => {
        "program t\nimplicit none\ninteger, dimension(2, 3) :: m\nm(2, 3) = 42\nprint *, m(2, 3)\nend program t\n",
        ["42"]
    };

    dimension_shorthand_real_vector => {
        "program t\nimplicit none\nreal :: v(5)\nv(5) = 9.0\nprint *, v(5)\nend program t\n",
        ["9"]
    };

    dimension_with_kind_attribute => {
        "program t\nimplicit none\ninteger(kind=4), dimension(3) :: arr\narr(2) = 77\nprint *, arr(2)\nend program t\n",
        ["77"]
    };

    dimension_trailing_bounds => {
        "program t\nimplicit none\ninteger :: grid(2, 2)\ngrid(1, 2) = 15\nprint *, grid(1, 2)\nend program t\n",
        ["15"]
    };

    dimension_parameter_bound => {
        "program t\nimplicit none\ninteger, parameter :: n = 3\ninteger, dimension(n) :: vec\nvec(n) = 100\nprint *, vec(n)\nend program t\n",
        ["100"]
    };

    logical_dimension_array => {
        "program t\nimplicit none\nlogical, dimension(3) :: flags\nflags(2) = .true.\nprint *, flags(2)\nend program t\n",
        ["true"]
    };

    // ── Parameter statements ───────────────────────────────────────────

    parameter_integer_expression => {
        "program t\nimplicit none\ninteger, parameter :: n = 2 + 3\nprint *, n\nend program t\n",
        ["5"]
    };

    parameter_real_expression => {
        "program t\nimplicit none\nreal, parameter :: tau = 2.0 * 3.14159\nprint *, tau\nend program t\n",
        ["6.28318"]
    };

    parameter_logical_false => {
        "program t\nimplicit none\nlogical, parameter :: nope = .false.\nprint *, nope\nend program t\n",
        ["false"]
    };

    parameter_character_literal => {
        "program t\nimplicit none\ncharacter(len=3), parameter :: tag = \"vyb\"\nprint *, tag\nend program t\n",
        ["vyb"]
    };

    multiple_parameters_same_statement => {
        "program t\nimplicit none\ninteger, parameter :: a = 1, b = 2, c = 3\nprint *, a + b + c\nend program t\n",
        ["6"]
    };

    parameter_used_in_array_bound => {
        "program t\nimplicit none\ninteger, parameter :: rows = 2, cols = 2\ninteger, dimension(rows, cols) :: mat\nmat(2, 1) = 8\nprint *, mat(2, 1)\nend program t\n",
        ["8"]
    };

    // ── Initialization ─────────────────────────────────────────────────

    init_integer_negative => {
        "program t\nimplicit none\ninteger :: x = -12\nprint *, x\nend program t\n",
        ["-12"]
    };

    init_real_from_integer => {
        "program t\nimplicit none\nreal :: r = 7\nprint *, r\nend program t\n",
        ["7"]
    };

    init_multiple_integers => {
        "program t\nimplicit none\ninteger :: p = 10, q = 20\nprint *, p + q\nend program t\n",
        ["30"]
    };

    init_character_from_parameter => {
        "program t\nimplicit none\ncharacter(len=6), parameter :: base = \"planet\"\ncharacter(len=6) :: word = base\nprint *, trim(word)\nend program t\n",
        ["planet"]
    };

    init_logical_from_comparison => {
        "program t\nimplicit none\nlogical :: gt = (5 > 3)\nprint *, gt\nend program t\n",
        ["true"]
    };

    // ── Implicit none programs ─────────────────────────────────────────

    implicit_none_integer_real => {
        "program t\nimplicit none\ninteger :: i = 4\nreal :: r = 2.0\nprint *, i + nint(r)\nend program t\n",
        ["6"]
    };

    implicit_none_logical_character => {
        "program t\nimplicit none\nlogical :: ok = .true.\ncharacter(len=2) :: ch = \"ok\"\nprint *, ok\nprint *, trim(ch)\nend program t\n",
        ["true", "ok"]
    };

    implicit_none_complex_print_real => {
        "program t\nimplicit none\ncomplex :: z = (3.0, 4.0)\nprint *, nint(real(z))\nend program t\n",
        ["3"]
    };

    implicit_none_double_precision => {
        "program t\nimplicit none\ndouble precision :: d = 2.5d0\nprint *, d\nend program t\n",
        ["2.5"]
    };

    implicit_none_mixed_declarations => {
        "program t\nimplicit none\ninteger :: a = 1\nreal :: b = 2.0\nlogical :: c = .true.\ncharacter(len=1) :: d = \"x\"\nprint *, a\nprint *, b\nprint *, c\nprint *, d\nend program t\n",
        ["1", "2", "true", "x"]
    };

    // ── Double precision ───────────────────────────────────────────────

    double_precision_init_literal => {
        "program t\ndouble precision :: d = 9.87654321d0\nprint *, d\nend program t\n",
        ["9.87654321"]
    };

    double_precision_assign_after_decl => {
        "program t\ndouble precision :: d\nd = 4.0d0\nprint *, d\nend program t\n",
        ["4"]
    };

    double_precision_parameter => {
        "program t\ndouble precision, parameter :: e = 2.718281828d0\nprint *, e\nend program t\n",
        ["2.718281828"]
    };

    double_precision_dimension_array => {
        "program t\ndouble precision, dimension(2) :: vals\nvals(1) = 1.5d0\nvals(2) = 2.5d0\nprint *, vals(2)\nend program t\n",
        ["2.5"]
    };

    // ── Complex declarations ───────────────────────────────────────────

    complex_literal_init_parts => {
        "program t\nimplicit none\ncomplex :: z = (6.0, -2.0)\nprint *, nint(real(z))\nprint *, nint(aimag(z))\nend program t\n",
        ["6", "-2"]
    };

    complex_kind_8_decl => {
        "program t\nimplicit none\ncomplex(kind=8) :: z = (1.0_8, 2.0_8)\nprint *, nint(real(z))\nprint *, nint(aimag(z))\nend program t\n",
        ["1", "2"]
    };

    double_complex_decl => {
        "program t\nimplicit none\ndouble complex :: z = (3.0d0, 4.0d0)\nprint *, nint(real(z))\nprint *, nint(aimag(z))\nend program t\n",
        ["3", "4"]
    };

    complex_parameter_const => {
        "program t\nimplicit none\ncomplex, parameter :: unit_i = (0.0, 1.0)\nprint *, nint(aimag(unit_i))\nend program t\n",
        ["1"]
    };

    complex_dimension_array => {
        "program t\nimplicit none\ncomplex, dimension(2) :: zs\nzs(1) = (1.0, 2.0)\nzs(2) = (3.0, 4.0)\nprint *, nint(real(zs(2)))\nprint *, nint(aimag(zs(2)))\nend program t\n",
        ["3", "4"]
    };

    complex_init_via_cmplx => {
        "program t\nimplicit none\ncomplex :: z\nz = cmplx(8.0, -1.0)\nprint *, nint(real(z))\nprint *, nint(aimag(z))\nend program t\n",
        ["8", "-1"]
    };
}
