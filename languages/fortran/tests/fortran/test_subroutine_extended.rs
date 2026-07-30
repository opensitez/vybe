//! Extended subroutine and function procedure coverage: internal procedures,
//! INTENT semantics, optional arguments, recursion, pure/elemental calls.

use super::helpers::run_prints;

fortran_cases! {
    // ── Internal subroutines ─────────────────────────────────────────

    internal_sub_calls_inner_print => {
        "program t\ncall outer()\ncontains\nsubroutine outer()\ncall inner(42)\nend subroutine outer\nsubroutine inner(v)\ninteger, intent(in) :: v\nprint *, v\nend subroutine inner\nend program t\n",
        ["42"]
    };

    internal_three_level_sum => {
        "program t\ncall level_a()\ncontains\nsubroutine level_a()\ncall level_b(1)\nend subroutine level_a\nsubroutine level_b(x)\ninteger, intent(in) :: x\ncall level_c(x, 2)\nend subroutine level_b\nsubroutine level_c(a, b)\ninteger, intent(in) :: a, b\nprint *, a + b\nend subroutine level_c\nend program t\n",
        ["3"]
    };

    internal_sub_invokes_local_function => {
        "program t\ncall report_square(6)\ncontains\nfunction square(n) result(r)\ninteger, intent(in) :: n\ninteger :: r\nr = n * n\nend function square\nsubroutine report_square(n)\ninteger, intent(in) :: n\nprint *, square(n)\nend subroutine report_square\nend program t\n",
        ["36"]
    };

    internal_two_subs_print_in_order => {
        "program t\ncall first()\ncall second()\ncontains\nsubroutine first()\nprint *, 1\nend subroutine first\nsubroutine second()\nprint *, 2\nend subroutine second\nend program t\n",
        ["1", "2"]
    };

    internal_sub_passes_result_to_inner => {
        "program t\ncall driver(4)\ncontains\nfunction triple(n) result(r)\ninteger, intent(in) :: n\ninteger :: r\nr = n * 3\nend function triple\nsubroutine driver(n)\ninteger, intent(in) :: n\ncall show(triple(n))\nend subroutine driver\nsubroutine show(v)\ninteger, intent(in) :: v\nprint *, v\nend subroutine show\nend program t\n",
        ["12"]
    };

    // ── Functions returning values ───────────────────────────────────

    function_real_average_two_values => {
        "program t\nprint *, avg2(6.0, 4.0)\ncontains\nreal function avg2(a, b)\nreal, intent(in) :: a, b\navg2 = (a + b) / 2.0\nend function avg2\nend program t\n",
        ["5"]
    };

    function_logical_is_positive_true => {
        "program t\nprint *, is_positive(7)\ncontains\nlogical function is_positive(n)\ninteger, intent(in) :: n\nis_positive = n > 0\nend function is_positive\nend program t\n",
        ["true"]
    };

    function_logical_is_positive_false => {
        "program t\nprint *, is_positive(-4)\ncontains\nlogical function is_positive(n)\ninteger, intent(in) :: n\nis_positive = n > 0\nend function is_positive\nend program t\n",
        ["false"]
    };

    function_integer_remainder_pair => {
        "program t\nprint *, rem_pair(17, 5)\ncontains\ninteger function rem_pair(a, b)\ninteger, intent(in) :: a, b\nrem_pair = mod(a, b)\nend function rem_pair\nend program t\n",
        ["2"]
    };

    function_result_variable_distinct => {
        "program t\nprint *, incr(9)\ncontains\nfunction incr(x) result(y)\ninteger, intent(in) :: x\ninteger :: y\ny = x + 1\nend function incr\nend program t\n",
        ["10"]
    };

    integer_prefix_function_difference => {
        "program t\nprint *, diff(15, 6)\ncontains\ninteger function diff(a, b)\ninteger, intent(in) :: a, b\ndiff = a - b\nend function diff\nend program t\n",
        ["9"]
    };

    // ── INTENT(in) ───────────────────────────────────────────────────

    intent_in_triple_add => {
        "program t\nprint *, add3(2, 3, 4)\ncontains\nfunction add3(a, b, c) result(s)\ninteger, intent(in) :: a, b, c\ninteger :: s\ns = a + b + c\nend function add3\nend program t\n",
        ["9"]
    };

    intent_in_product_pair => {
        "program t\nprint *, mul2(6, 7)\ncontains\nfunction mul2(x, y) result(p)\ninteger, intent(in) :: x, y\ninteger :: p\np = x * y\nend function mul2\nend program t\n",
        ["42"]
    };

    intent_in_array_minimum => {
        "program t\ninteger :: v(4)\nv = [8, 3, 11, 5]\nprint *, arr_min(v, 4)\ncontains\nfunction arr_min(a, n) result(m)\ninteger, intent(in) :: a(n), n\ninteger :: m, i\nm = a(1)\ndo i = 2, n\nif (a(i) < m) m = a(i)\nend do\nend function arr_min\nend program t\n",
        ["3"]
    };

    intent_in_dot_product_custom => {
        "program t\ninteger :: x(2), y(2)\nx = [1, 2]\ny = [3, 4]\nprint *, dot2(x, y, 2)\ncontains\nfunction dot2(u, v, n) result(s)\ninteger, intent(in) :: u(n), v(n), n\ninteger :: s, i\ns = 0\ndo i = 1, n\ns = s + u(i) * v(i)\nend do\nend function dot2\nend program t\n",
        ["11"]
    };

    // ── INTENT(out) ──────────────────────────────────────────────────

    intent_out_pair_scalars => {
        "program t\ninteger :: lo, hi\ncall bounds(lo, hi)\nprint *, lo\nprint *, hi\ncontains\nsubroutine bounds(a, b)\ninteger, intent(out) :: a, b\na = 3\nb = 11\nend subroutine bounds\nend program t\n",
        ["3", "11"]
    };

    intent_out_single_assign => {
        "program t\ninteger :: n\ncall fill(n)\nprint *, n\ncontains\nsubroutine fill(x)\ninteger, intent(out) :: x\nx = 99\nend subroutine fill\nend program t\n",
        ["99"]
    };

    intent_out_zero_array => {
        "program t\ninteger :: a(3)\ncall zero_fill(a)\nprint *, sum(a)\ncontains\nsubroutine zero_fill(arr)\ninteger, intent(out) :: arr(3)\narr = 0\nend subroutine zero_fill\nend program t\n",
        ["0"]
    };

    intent_out_two_outputs_from_sub => {
        "program t\ninteger :: p, q\ncall split_sum(10, p, q)\nprint *, p\nprint *, q\ncontains\nsubroutine split_sum(n, half, rest)\ninteger, intent(in) :: n\ninteger, intent(out) :: half, rest\nhalf = n / 2\nrest = n - half\nend subroutine split_sum\nend program t\n",
        ["5", "5"]
    };

    // ── INTENT(inout) ────────────────────────────────────────────────

    intent_inout_add_five => {
        "program t\ninteger :: x\nx = 10\ncall add_five(x)\nprint *, x\ncontains\nsubroutine add_five(n)\ninteger, intent(inout) :: n\nn = n + 5\nend subroutine add_five\nend program t\n",
        ["15"]
    };

    intent_inout_halve_integer => {
        "program t\ninteger :: x\nx = 20\ncall halve(x)\nprint *, x\ncontains\nsubroutine halve(n)\ninteger, intent(inout) :: n\nn = n / 2\nend subroutine halve\nend program t\n",
        ["10"]
    };

    intent_inout_swap_scalars => {
        "program t\ninteger :: a, b\na = 3\nb = 7\ncall swap_int(a, b)\nprint *, a\nprint *, b\ncontains\nsubroutine swap_int(x, y)\ninteger, intent(inout) :: x, y\ninteger :: t\nt = x\nx = y\ny = t\nend subroutine swap_int\nend program t\n",
        ["7", "3"]
    };

    intent_inout_bump_array_sum => {
        "program t\ninteger :: a(3)\na = [1, 2, 3]\ncall bump_all(a)\nprint *, sum(a)\ncontains\nsubroutine bump_all(v)\ninteger, intent(inout) :: v(3)\nv = v + 1\nend subroutine bump_all\nend program t\n",
        ["9"]
    };

    // ── Optional arguments ─────────────────────────────────────────

    optional_addend_missing_and_present => {
        "program t\nprint *, with_addend(5)\nprint *, with_addend(5, 3)\ncontains\nfunction with_addend(x, y) result(r)\ninteger, intent(in) :: x\ninteger, intent(in), optional :: y\ninteger :: r\nif (present(y)) then\nr = x + y\nelse\nr = x\nend if\nend function with_addend\nend program t\n",
        ["5", "8"]
    };

    optional_multiplier_defaults_to_one => {
        "program t\nprint *, scale_val(6)\nprint *, scale_val(6, 4)\ncontains\nfunction scale_val(x, factor) result(r)\ninteger, intent(in) :: x\ninteger, intent(in), optional :: factor\ninteger :: r\nif (present(factor)) then\nr = x * factor\nelse\nr = x\nend if\nend function scale_val\nend program t\n",
        ["6", "24"]
    };

    optional_subroutine_extra_term => {
        "program t\ncall accumulate(4)\ncall accumulate(4, 6)\ncontains\nsubroutine accumulate(base, extra)\ninteger, intent(in) :: base\ninteger, intent(in), optional :: extra\ninteger :: total\ntotal = base\nif (present(extra)) total = total + extra\nprint *, total\nend subroutine accumulate\nend program t\n",
        ["4", "10"]
    };

    optional_title_prefix_when_present => {
        "program t\ncall greet_opt('Ann')\ncall greet_opt('Ann', 'Ms.')\ncontains\nsubroutine greet_opt(name, title)\ncharacter(len=*), intent(in) :: name\ncharacter(len=*), intent(in), optional :: title\nif (present(title)) then\nprint *, trim(title) // ' ' // trim(name)\nelse\nprint *, trim(name)\nend if\nend subroutine greet_opt\nend program t\n",
        ["Ann", "Ms. Ann"]
    };

    // ── Recursive procedures ─────────────────────────────────────────

    recursive_fibonacci_ten => {
        "program t\nprint *, fib(10)\ncontains\nrecursive function fib(n) result(r)\ninteger, intent(in) :: n\ninteger :: r\nif (n <= 1) then\nr = n\nelse\nr = fib(n - 1) + fib(n - 2)\nend if\nend function fib\nend program t\n",
        ["55"]
    };

    recursive_sum_one_to_five => {
        "program t\nprint *, series_sum(5)\ncontains\nrecursive function series_sum(n) result(s)\ninteger, intent(in) :: n\ninteger :: s\nif (n <= 0) then\ns = 0\nelse\ns = n + series_sum(n - 1)\nend if\nend function series_sum\nend program t\n",
        ["15"]
    };

    recursive_factorial_four => {
        "program t\nprint *, fact4(4)\ncontains\nrecursive function fact4(n) result(r)\ninteger, intent(in) :: n\ninteger :: r\nif (n <= 1) then\nr = 1\nelse\nr = n * fact4(n - 1)\nend if\nend function fact4\nend program t\n",
        ["24"]
    };

    recursive_countdown_subroutine => {
        "program t\ncall count_down(3)\ncontains\nrecursive subroutine count_down(n)\ninteger, intent(in) :: n\nprint *, n\nif (n > 1) call count_down(n - 1)\nend subroutine count_down\nend program t\n",
        ["3", "2", "1"]
    };

    recursive_digit_sum_four_digits => {
        "program t\nprint *, digit_sum(9876)\ncontains\nrecursive function digit_sum(n) result(s)\ninteger, intent(in) :: n\ninteger :: s\nif (n < 10) then\ns = n\nelse\ns = mod(n, 10) + digit_sum(n / 10)\nend if\nend function digit_sum\nend program t\n",
        ["30"]
    };

    recursive_gcd_six_and_four => {
        "program t\nprint *, my_gcd(6, 4)\ncontains\nrecursive function my_gcd(a, b) result(g)\ninteger, intent(in) :: a, b\ninteger :: g\nif (b == 0) then\ng = a\nelse\ng = my_gcd(b, mod(a, b))\nend if\nend function my_gcd\nend program t\n",
        ["2"]
    };

    recursive_power_two_to_fifth => {
        "program t\nprint *, ipow(2, 5)\ncontains\nrecursive function ipow(base, exp) result(r)\ninteger, intent(in) :: base, exp\ninteger :: r\nif (exp == 0) then\nr = 1\nelse\nr = base * ipow(base, exp - 1)\nend if\nend function ipow\nend program t\n",
        ["32"]
    };

    // ── Pure functions printing results ──────────────────────────────

    pure_square_of_six => {
        "program t\nprint *, psquare(6)\ncontains\npure function psquare(x) result(r)\ninteger, intent(in) :: x\ninteger :: r\nr = x * x\nend function psquare\nend program t\n",
        ["36"]
    };

    pure_cube_of_three => {
        "program t\nprint *, pcube(3)\ncontains\npure function pcube(x) result(r)\ninteger, intent(in) :: x\ninteger :: r\nr = x * x * x\nend function pcube\nend program t\n",
        ["27"]
    };

    pure_integer_max_pair => {
        "program t\nprint *, pmax(12, 9)\ncontains\npure function pmax(a, b) result(m)\ninteger, intent(in) :: a, b\ninteger :: m\nif (a >= b) then\nm = a\nelse\nm = b\nend if\nend function pmax\nend program t\n",
        ["12"]
    };

    pure_add_three_terms => {
        "program t\nprint *, padd3(2, 3, 4)\ncontains\npure function padd3(a, b, c) result(s)\ninteger, intent(in) :: a, b, c\ninteger :: s\ns = a + b + c\nend function padd3\nend program t\n",
        ["9"]
    };

    pure_subroutine_doubles_inout => {
        "program t\ninteger :: n\nn = 11\ncall pdbl(n)\nprint *, n\ncontains\npure subroutine pdbl(x)\ninteger, intent(inout) :: x\nx = x * 2\nend subroutine pdbl\nend program t\n",
        ["22"]
    };

    // ── Elemental function calls on arrays ───────────────────────────

    elemental_double_array_sum => {
        "program t\ninteger :: a(4), b(4)\na = [1, 2, 3, 4]\nb = edouble(a)\nprint *, sum(b)\ncontains\nelemental function edouble(x) result(r)\ninteger, intent(in) :: x\ninteger :: r\nr = x * 2\nend function edouble\nend program t\n",
        ["20"]
    };

    elemental_square_first_element => {
        "program t\ninteger :: a(3), b(3)\na = [2, 3, 5]\nb = esquare(a)\nprint *, b(1)\ncontains\nelemental function esquare(x) result(r)\ninteger, intent(in) :: x\ninteger :: r\nr = x * x\nend function esquare\nend program t\n",
        ["4"]
    };

    elemental_add_ten_second_element => {
        "program t\ninteger :: a(3), b(3)\na = [1, 2, 3]\nb = eplus10(a)\nprint *, b(2)\ncontains\nelemental function eplus10(x) result(r)\ninteger, intent(in) :: x\ninteger :: r\nr = x + 10\nend function eplus10\nend program t\n",
        ["12"]
    };

    elemental_negate_array_sum => {
        "program t\ninteger :: a(3), b(3)\na = [1, 2, 3]\nb = eneg(a)\nprint *, sum(b)\ncontains\nelemental function eneg(x) result(r)\ninteger, intent(in) :: x\ninteger :: r\nr = -x\nend function eneg\nend program t\n",
        ["-6"]
    };

    elemental_scale_by_three_sum => {
        "program t\ninteger :: a(3), b(3)\na = [1, 2, 3]\nb = escale3(a)\nprint *, sum(b)\ncontains\nelemental function escale3(x) result(r)\ninteger, intent(in) :: x\ninteger :: r\nr = x * 3\nend function escale3\nend program t\n",
        ["18"]
    };

    elemental_real_sqrt_second => {
        "program t\nreal :: a(3), b(3)\na = [4.0, 9.0, 16.0]\nb = esqrt(a)\nprint *, b(2)\ncontains\nelemental function esqrt(x) result(r)\nreal, intent(in) :: x\nreal :: r\nr = sqrt(x)\nend function esqrt\nend program t\n",
        ["3"]
    };

    elemental_subroutine_negate_inout => {
        "program t\ninteger :: a(3)\na = [5, -2, 7]\ncall enegate(a)\nprint *, sum(a)\ncontains\nelemental subroutine enegate(x)\ninteger, intent(inout) :: x\nx = -x\nend subroutine enegate\nend program t\n",
        ["-10"]
    };

    elemental_pure_abs_sum => {
        "program t\ninteger :: a(4), b(4)\na = [-1, 2, -3, 4]\nb = eabsval(a)\nprint *, sum(b)\ncontains\nelemental function eabsval(x) result(r)\ninteger, intent(in) :: x\ninteger :: r\nif (x < 0) then\nr = -x\nelse\nr = x\nend if\nend function eabsval\nend program t\n",
        ["10"]
    };

    // ── Mixed procedure interactions ─────────────────────────────────

    sub_calls_func_then_prints => {
        "program t\ncall emit_square(5)\ncontains\nfunction sq(n) result(r)\ninteger, intent(in) :: n\ninteger :: r\nr = n * n\nend function sq\nsubroutine emit_square(n)\ninteger, intent(in) :: n\nprint *, sq(n)\nend subroutine emit_square\nend program t\n",
        ["25"]
    };

    nested_function_call_chain => {
        "program t\nprint *, twice(twice(3))\ncontains\nfunction twice(n) result(r)\ninteger, intent(in) :: n\ninteger :: r\nr = n * 2\nend function twice\nend program t\n",
        ["12"]
    };

    func_uses_pure_elemental_on_scalar => {
        "program t\nprint *, eincr(8)\ncontains\nelemental function eincr(x) result(r)\ninteger, intent(in) :: x\ninteger :: r\nr = x + 1\nend function eincr\nend program t\n",
        ["9"]
    };

    internal_sub_calls_pure_function => {
        "program t\ncall show_diff(9, 4)\ncontains\npure function pdiff(a, b) result(d)\ninteger, intent(in) :: a, b\ninteger :: d\nd = a - b\nend function pdiff\nsubroutine show_diff(x, y)\ninteger, intent(in) :: x, y\nprint *, pdiff(x, y)\nend subroutine show_diff\nend program t\n",
        ["5"]
    };
}

// ── Compile-only procedure shapes ───────────────────────────────────

#[test]
fn compile_mutual_recursive_even_odd() {
    let out = run_prints(
        r#"
program t
    print *, is_even(6)
contains
    recursive function is_even(n) result(b)
        integer, intent(in) :: n
        logical :: b
        if (n == 0) then
            b = .true.
        else
            b = is_odd(n - 1)
        end if
    end function is_even

    recursive function is_odd(n) result(b)
        integer, intent(in) :: n
        logical :: b
        if (n == 0) then
            b = .false.
        else
            b = is_even(n - 1)
        end if
    end function is_odd
end program t
"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn compile_optional_three_parameter_subroutine() {
    let out = run_prints(
        r#"
program t
    call tagged(1)
    call tagged(1, 2)
    call tagged(1, 2, 3)
contains
    subroutine tagged(a, b, c)
        integer, intent(in) :: a
        integer, intent(in), optional :: b, c
        integer :: total
        total = a
        if (present(b)) total = total + b
        if (present(c)) total = total + c
        print *, total
    end subroutine tagged
end program t
"#,
    );
    assert_eq!(out, vec!["1", "3", "6"]);
}

#[test]
fn compile_nested_contains_many_procedures() {
    let out = run_prints(
        r#"
program t
    print *, f1(1) + f2(2) + f3(3)
contains
    function f1(x) result(r)
        integer, intent(in) :: x
        integer :: r
        r = x
    end function f1
    function f2(x) result(r)
        integer, intent(in) :: x
        integer :: r
        r = x
    end function f2
    function f3(x) result(r)
        integer, intent(in) :: x
        integer :: r
        r = x
    end function f3
    subroutine noop()
end subroutine noop
end program t
"#,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn compile_pure_elemental_function_signature() {
    let out = run_prints(
        r#"
program t
    integer :: a(2) = [1, 2]
    print *, blend(a, 0)
contains
    elemental function blend(x, bias) result(r)
        integer, intent(in) :: x, bias
        integer :: r
        r = x + bias
end function blend
end program t
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn compile_intent_mixed_subroutine_signature() {
    let out = run_prints(
        r#"
program t
    integer :: x, y(2)
    y = [4, 5]
    call mixer(3, x, y)
contains
    subroutine mixer(n, s, arr)
        integer, intent(in) :: n
        integer, intent(out) :: s
        integer, intent(inout) :: arr(2)
        s = n
        arr = arr + n
    end subroutine mixer
end program t
"#,
    );
    assert_eq!(out, vec!["3", "15"]);
}

#[test]
fn compile_recursive_function_with_result_clause() {
    let out = run_prints(
        r#"
program t
    print *, len_num(12345)
contains
    recursive function len_num(n) result(d)
        integer, intent(in) :: n
        integer :: d
        if (n < 10) then
            d = 1
        else
            d = 1 + len_num(n / 10)
        end if
    end function len_num
end program t
"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn mutual_recursive_even_odd_runtime_paths() {
    let out = run_prints(
        r#"
program t
    print *, is_even(7)
    print *, is_odd(8)
contains
    recursive function is_even(n) result(b)
        integer, intent(in) :: n
        logical :: b
        if (n == 0) then
            b = .true.
        else
            b = is_odd(n - 1)
        end if
    end function is_even

    recursive function is_odd(n) result(b)
        integer, intent(in) :: n
        logical :: b
        if (n == 0) then
            b = .false.
        else
            b = is_even(n - 1)
        end if
    end function is_odd
end program t
"#,
    );
    assert_eq!(out, vec!["False", "False"]);
}

#[test]
fn optional_parameter_coverage_for_third_argument_runtime() {
    let out = run_prints(
        r#"
program t
    call tagged(1)
    call tagged(1, 2)
    call tagged(1, 2, 3)
contains
    subroutine tagged(a, b, c)
        integer, intent(in) :: a
        integer, intent(in), optional :: b, c
        integer :: total
        total = a
        if (present(b)) then
            total = total + b
        end if
        if (present(c)) then
            total = total + c
        end if
        print *, total
    end subroutine tagged
end program t
"#,
    );
    assert_eq!(out, vec!["1", "3", "6"]);
}

#[test]
fn mixed_intent_signature_runtime_results() {
    let out = run_prints(
        r#"
program t
    integer :: x
    integer :: y(2)
    y = [4, 5]
    call mixer(3, x, y)
    print *, x
    print *, sum(y)
contains
    subroutine mixer(n, s, arr)
        integer, intent(in) :: n
        integer, intent(out) :: s
        integer, intent(inout) :: arr(2)
        s = n
        arr = arr + n
    end subroutine mixer
end program t
"#,
    );
    assert_eq!(out, vec!["3", "15"]);
}

#[test]
fn elemental_signature_for_arrays_runtime() {
    let out = run_prints(
        r#"
program t
    integer :: a(2) = [1, 2]
    print *, sum(blend(a, 0))
    print *, sum(blend(a, 4))
contains
    elemental function blend(x, bias) result(r)
        integer, intent(in) :: x, bias
        integer :: r
        r = x + bias
    end function blend
end program t
"#,
    );
    assert_eq!(out, vec!["3", "11"]);
}
