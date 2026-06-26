//! Procedure attribute semantics: RECURSIVE traces, BIND(C), ELEMENTAL arrays,
//! PURE violations (compile-only), internal nesting, module vs external procedures.

use super::helpers::compile_ok;

fortran_cases! {
    recursive_factorial_fibonacci_trace_prints => {
        "program t\nprint *, fact_trace(4)\nprint *, fib_trace(6)\ncontains\nrecursive function fact_trace(n) result(r)\ninteger, intent(in) :: n\ninteger :: r\nprint *, n\nif (n <= 1) then\nr = 1\nelse\nr = n * fact_trace(n - 1)\nend if\nend function fact_trace\nrecursive function fib_trace(n) result(r)\ninteger, intent(in) :: n\ninteger :: r\nif (n <= 1) then\nr = n\nelse\nr = fib_trace(n - 1) + fib_trace(n - 2)\nend if\nprint *, r\nend function fib_trace\nend program t\n",
        ["4", "3", "2", "1", "8", "5", "3", "2", "1", "1", "0", "8"]
    };

    elemental_custom_abs_diff_on_arrays => {
        "program t\ninteger :: a(3), b(3), c(3)\na = [5, 1, 9]\nb = [2, 4, 3]\nc = abs_diff(a, b)\nprint *, c(1)\nprint *, sum(c)\ncontains\nelemental function abs_diff(x, y) result(d)\ninteger, intent(in) :: x, y\ninteger :: d\nd = abs(x - y)\nend function abs_diff\nend program t\n",
        ["3", "10"]
    };

    internal_nested_contains_function_chain => {
        "program t\ncall outer_driver()\ncontains\nsubroutine outer_driver()\ncall middle_layer(7)\ncontains\nsubroutine middle_layer(v)\ninteger, intent(in) :: v\nprint *, inner_fn(v)\ncontains\nfunction inner_fn(n) result(r)\ninteger, intent(in) :: n\ninteger :: r\nr = n + 3\nend function inner_fn\nend subroutine middle_layer\nend subroutine outer_driver\nend program t\n",
        ["10"]
    };

    module_procedure_vs_external_same_shape => {
        "module attr_mod\nimplicit none\ninterface add_pair\nmodule procedure mod_add\nend interface\ncontains\nfunction mod_add(a, b) result(r)\ninteger, intent(in) :: a, b\ninteger :: r\nr = a + b\nend function mod_add\nend module attr_mod\nfunction ext_add(a, b) result(r)\ninteger, intent(in) :: a, b\ninteger :: r\nr = a * b\nend function ext_add\nprogram t\nuse attr_mod\nprint *, add_pair(3, 4)\nprint *, ext_add(3, 4)\nend program t\n",
        ["7", "12"]
    };
}

#[test]
fn compile_bind_c_fortran_function() {
    compile_ok(
        r#"
module c_proc
    use iso_c_binding
    implicit none
contains
    function add_c(a, b) bind(c, name='add_c') result(r)
        integer(c_int), intent(in), value :: a, b
        integer(c_int) :: r
        r = a + b
    end function add_c
end module c_proc

program t
    use c_proc
    print *, "ok"
end program t
"#,
    );
}

#[test]
fn compile_pure_function_with_print_violation() {
    compile_ok(
        r#"
program t
    print *, bad_pure(5)
contains
    pure function bad_pure(x) result(r)
        integer, intent(in) :: x
        integer :: r
        print *, x
        r = x * 2
    end function bad_pure
end program t
"#,
    );
}
