//! Extended submodule coverage: parent parameters and variables, INTENT(out),
//! logical/character/kind-8 results, optional args, and nested interface-only
//! child submodules. Distinct from `test_submodules_advanced.rs`.

use super::helpers::run_prints;

fortran_cases! {
    // ── Runnable: submodule units compile; program uses parent symbols ─

    submodule_single_unit_parent_parameter_visible => {
        "module host\nimplicit none\ninteger, parameter :: TAG = 99\ninterface\nmodule function bump(x) result(r)\ninteger, intent(in) :: x\ninteger :: r\nend function bump\nend interface\nend module host\nsubmodule (host) host_impl\ncontains\nmodule function bump(x) result(r)\ninteger, intent(in) :: x\ninteger :: r\nr = x + TAG\nend function bump\nend submodule host_impl\nprogram t\nuse host\nprint *, TAG\nend program t\n",
        ["99"]
    };

    submodule_two_units_share_parent_constant => {
        "module pair_iface\nimplicit none\ninteger, parameter :: BASE = 3\ninterface\nmodule function add1(x) result(r)\ninteger, intent(in) :: x\ninteger :: r\nend function add1\nmodule function add2(x) result(r)\ninteger, intent(in) :: x\ninteger :: r\nend function add2\nend interface\nend module pair_iface\nsubmodule (pair_iface) pair_a\ncontains\nmodule function add1(x) result(r)\ninteger, intent(in) :: x\ninteger :: r\nr = x + BASE\nend function add1\nend submodule pair_a\nsubmodule (pair_iface) pair_b\ncontains\nmodule function add2(x) result(r)\ninteger, intent(in) :: x\ninteger :: r\nr = x + BASE + BASE\nend function add2\nend submodule pair_b\nprogram t\nuse pair_iface\nprint *, BASE\nend program t\n",
        ["3"]
    };

    submodule_nested_units_parent_anchor_constant => {
        "module anchor_iface\nimplicit none\ninteger, parameter :: ANCHOR = 77\ninterface\nmodule function lift(x) result(r)\ninteger, intent(in) :: x\ninteger :: r\nend function lift\nend interface\nend module anchor_iface\nsubmodule (anchor_iface) mid_iface\ninterface\nmodule function helper(x) result(r)\ninteger, intent(in) :: x\ninteger :: r\nend function helper\nend interface\nend submodule mid_iface\nsubmodule (anchor_iface:mid_iface) bot_impl\ncontains\nmodule function lift(x) result(r)\ninteger, intent(in) :: x\ninteger :: r\nr = x + ANCHOR\nend function lift\nmodule function helper(x) result(r)\ninteger, intent(in) :: x\ninteger :: r\nr = x + 1\nend function helper\nend submodule bot_impl\nprogram t\nuse anchor_iface\nprint *, ANCHOR\nend program t\n",
        ["77"]
    };

    submodule_use_rename_binding_from_submodule => {
        "module scale_iface\nimplicit none\ninterface\nmodule function scale(x) result(r)\ninteger, intent(in) :: x\ninteger :: r\nend function scale\nend interface\nend module scale_iface\nsubmodule (scale_iface) scale_impl\ncontains\nmodule function scale(x) result(r)\ninteger, intent(in) :: x\ninteger :: r\nr = x * 3\nend function scale\nend submodule scale_impl\nprogram t\nuse scale_iface, only: triple => scale\nprint *, triple(4)\nend program t\n",
        ["12"]
    };

    submodule_parent_contained_helper_callable_in_submodule => {
        "module helper_iface\nimplicit none\ninterface\nmodule function boosted(x) result(r)\ninteger, intent(in) :: x\ninteger :: r\nend function boosted\nend interface\ncontains\ninteger function local_offset()\nlocal_offset = 2\nend function local_offset\nend module helper_iface\nsubmodule (helper_iface) helper_impl\ncontains\nmodule function boosted(x) result(r)\ninteger, intent(in) :: x\ninteger :: r\nr = x + local_offset()\nend function boosted\nend submodule helper_impl\nprogram t\nuse helper_iface\nprint *, boosted(5)\nend program t\n",
        ["7"]
    };

    submodule_parent_type_ctor_without_calling_iface => {
        "module geom_iface\nimplicit none\ntype :: Point\nreal :: x, y\nend type Point\ninterface\nmodule function distance(a, b) result(d)\ntype(Point), intent(in) :: a, b\nreal :: d\nend function distance\nend interface\nend module geom_iface\nsubmodule (geom_iface) geom_impl\ncontains\nmodule function distance(a, b) result(d)\ntype(Point), intent(in) :: a, b\nreal :: d\nd = sqrt((a%x - b%x)**2 + (a%y - b%y)**2)\nend function distance\nend submodule geom_impl\nprogram t\nuse geom_iface\ntype(Point) :: p\np%x = 3.0\np%y = 4.0\nprint *, int(p%x)\nend program t\n",
        ["3"]
    };

    submodule_parent_var_initial_before_sub_call => {
        "module bank_iface\nimplicit none\ninteger :: balance = 100\ninterface\nmodule subroutine deposit(amt)\ninteger, intent(in) :: amt\nend subroutine deposit\nend interface\nend module bank_iface\nsubmodule (bank_iface) bank_impl\ncontains\nmodule subroutine deposit(amt)\ninteger, intent(in) :: amt\nbalance = balance + amt\nend subroutine deposit\nend submodule bank_impl\nprogram t\nuse bank_iface\nprint *, balance\nend program t\n",
        ["100"]
    };

    submodule_generic_iface_present_program_sums_array => {
        "module generic_iface\nimplicit none\ninterface norm\nmodule function norm_real(a) result(r)\nreal, intent(in) :: a(:)\nreal :: r\nend function norm_real\nend interface norm\nend module generic_iface\nsubmodule (generic_iface) generic_impl\ncontains\nmodule function norm_real(a) result(r)\nreal, intent(in) :: a(:)\nreal :: r\nr = sqrt(sum(a**2))\nend function norm_real\nend submodule generic_impl\nprogram t\nuse generic_iface\nreal :: v(3) = [3.0, 4.0, 0.0]\nprint *, int(sum(v))\nend program t\n",
        ["7"]
    };
}

// ── Parent symbols in submodule implementation (compile-only) ────────

#[test]
fn submodule_uses_parent_parameter_in_body() {
    let out = run_prints(
        r#"
module scale_iface
    implicit none
    integer, parameter :: SCALE = 10
    interface
        module function scaled(x) result(r)
            integer, intent(in) :: x
            integer :: r
        end function scaled
    end interface
end module scale_iface

submodule (scale_iface) scale_impl
    implicit none
contains
    module function scaled(x) result(r)
        integer, intent(in) :: x
        integer :: r
        r = x * SCALE
    end function scaled
end submodule scale_impl

program t
    use scale_iface
    print *, scaled(4)
end program t
"#,
    );
    assert_eq!(out, vec!["40"]);
}

#[test]
fn submodule_mutates_parent_public_variable() {
    let out = run_prints(
        r#"
module bank_iface
    implicit none
    integer :: balance = 100
    interface
        module subroutine deposit(amt)
            integer, intent(in) :: amt
        end subroutine deposit
    end interface
end module bank_iface

submodule (bank_iface) bank_impl
contains
    module subroutine deposit(amt)
        integer, intent(in) :: amt
        balance = balance + amt
    end subroutine deposit
end submodule bank_impl

program t
    use bank_iface
    call deposit(50)
    print *, balance
end program t
"#,
    );
    assert_eq!(out, vec!["150"]);
}

#[test]
fn submodule_reads_parent_module_variable() {
    let out = run_prints(
        r#"
module mult_iface
    implicit none
    integer :: factor = 7
    interface
        module function times_factor(n) result(r)
            integer, intent(in) :: n
            integer :: r
        end function times_factor
    end interface
end module mult_iface

submodule (mult_iface) mult_impl
contains
    module function times_factor(n) result(r)
        integer, intent(in) :: n
        integer :: r
        r = n * factor
    end function times_factor
end submodule mult_impl

program t
    use mult_iface
    print *, times_factor(6)
end program t
"#,
    );
    assert_eq!(out, vec!["42"]);
}

// ── INTENT, optional, and result types (compile-only) ───────────────

#[test]
fn submodule_intent_out_fills_vector() {
    let out = run_prints(
        r#"
module fill_iface
    implicit none
    interface
        module subroutine fill_seq(v, n)
            integer, intent(out) :: v(:)
            integer, intent(in) :: n
        end subroutine fill_seq
    end interface
end module fill_iface

submodule (fill_iface) fill_impl
contains
    module subroutine fill_seq(v, n)
        integer, intent(out) :: v(:)
        integer, intent(in) :: n
        integer :: i
        do i = 1, n
            v(i) = i * 2
        end do
    end subroutine fill_seq
end submodule fill_impl

program t
    use fill_iface
    integer :: a(3)
    call fill_seq(a, 3)
    print *, a(2)
end program t
"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn submodule_three_intent_in_arguments() {
    let out = run_prints(
        r#"
module sum3_iface
    implicit none
    interface
        module function sum3(a, b, c) result(s)
            integer, intent(in) :: a, b, c
            integer :: s
        end function sum3
    end interface
end module sum3_iface

submodule (sum3_iface) sum3_impl
contains
    module function sum3(a, b, c) result(s)
        integer, intent(in) :: a, b, c
        integer :: s
        s = a + b + c
    end function sum3
end submodule sum3_impl

program t
    use sum3_iface
    print *, sum3(2, 3, 4)
end program t
"#,
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn submodule_optional_second_argument_default() {
    let out = run_prints(
        r#"
module opt_iface
    implicit none
    interface
        module function bump(n, step) result(r)
            integer, intent(in) :: n
            integer, optional, intent(in) :: step
            integer :: r
        end function bump
    end interface
end module opt_iface

submodule (opt_iface) opt_impl
contains
    module function bump(n, step) result(r)
        integer, intent(in) :: n
        integer, optional, intent(in) :: step
        integer :: r
        if (present(step)) then
            r = n + step
        else
            r = n + 1
        end if
    end function bump
end submodule opt_impl

program t
    use opt_iface
    print *, bump(5)
    print *, bump(5, 3)
end program t
"#,
    );
    assert_eq!(out, vec!["6", "8"]);
}

#[test]
fn submodule_logical_result_function() {
    let out = run_prints(
        r#"
module even_iface
    implicit none
    interface
        module function is_even(n) result(flag)
            integer, intent(in) :: n
            logical :: flag
        end function is_even
    end interface
end module even_iface

submodule (even_iface) even_impl
contains
    module function is_even(n) result(flag)
        integer, intent(in) :: n
        logical :: flag
        flag = mod(n, 2) == 0
    end function is_even
end submodule even_impl

program t
    use even_iface
    print *, is_even(8)
    print *, is_even(7)
end program t
"#,
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn submodule_character_scalar_result() {
    let out = run_prints(
        r#"
module char_iface
    implicit none
    interface
        module function first_char(s) result(c)
            character(len=*), intent(in) :: s
            character(len=1) :: c
        end function first_char
    end interface
end module char_iface

submodule (char_iface) char_impl
contains
    module function first_char(s) result(c)
        character(len=*), intent(in) :: s
        character(len=1) :: c
        c = s(1:1)
    end function first_char
end submodule char_impl

program t
    use char_iface
    print *, first_char('delta')
end program t
"#,
    );
    assert_eq!(out, vec!["d"]);
}

#[test]
fn submodule_kind8_real_result() {
    let out = run_prints(
        r#"
module dbl_iface
    implicit none
    interface
        module function halve(x) result(r)
            real(kind=8), intent(in) :: x
            real(kind=8) :: r
        end function halve
    end interface
end module dbl_iface

submodule (dbl_iface) dbl_impl
contains
    module function halve(x) result(r)
        real(kind=8), intent(in) :: x
        real(kind=8) :: r
        r = x / 2.0d0
    end function halve
end submodule dbl_impl

program t
    use dbl_iface
    print *, int(halve(9.0d0))
end program t
"#,
    );
    assert_eq!(out, vec!["4"]);
}

// ── Submodule-local state and array logic (compile-only) ─────────────

#[test]
fn submodule_array_reduction_over_argument() {
    let out = run_prints(
        r#"
module peak_iface
    implicit none
    interface
        module function peak(v) result(m)
            integer, intent(in) :: v(:)
            integer :: m
        end function peak
    end interface
end module peak_iface

submodule (peak_iface) peak_impl
contains
    module function peak(v) result(m)
        integer, intent(in) :: v(:)
        integer :: m
        m = v(1)
        if (v(2) > m) m = v(2)
        if (v(3) > m) m = v(3)
    end function peak
end submodule peak_impl

program t
    use peak_iface
    integer :: data(3)
    data = [3, 11, 7]
    print *, peak(data)
end program t
"#,
    );
    assert_eq!(out, vec!["11"]);
}

#[test]
fn submodule_reset_subroutine_clears_state() {
    let out = run_prints(
        r#"
module acc_iface
    implicit none
    interface
        module subroutine add_one()
        end subroutine add_one
        module subroutine reset_acc()
        end subroutine reset_acc
        module function read_acc() result(n)
            integer :: n
        end function read_acc
    end interface
end module acc_iface

submodule (acc_iface) acc_impl
    integer :: total = 0
contains
    module subroutine add_one()
        total = total + 1
    end subroutine add_one

    module subroutine reset_acc()
        total = 0
    end subroutine reset_acc

    module function read_acc() result(n)
        integer :: n
        n = total
    end function read_acc
end submodule acc_impl

program t
    use acc_iface
    call add_one()
    call add_one()
    call reset_acc()
    call add_one()
    print *, read_acc()
end program t
"#,
    );
    assert_eq!(out, vec!["1"]);
}

// ── Nested interface-only child; PURE; INTENT(inout); allocatable ───

#[test]
fn submodule_child_interface_grandchild_implements() {
    let out = run_prints(
        r#"
module top_iface
    implicit none
    interface
        module function top_val() result(r)
            integer :: r
        end function top_val
    end interface
end module top_iface

submodule (top_iface) mid_iface
    implicit none
    interface
        module function mid_val() result(r)
            integer :: r
        end function mid_val
    end interface
end submodule mid_iface

submodule (top_iface:mid_iface) bot_impl
    implicit none
contains
    module function top_val() result(r)
        integer :: r
        r = 10
    end function top_val

    module function mid_val() result(r)
        integer :: r
        r = 20
    end function mid_val
end submodule bot_impl

program t
    use top_iface
    print *, top_val()
end program t
"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn submodule_pure_module_procedure() {
    let out = run_prints(
        r#"
module pure_iface
    implicit none
    interface
        module function pure_add(a, b) result(r)
            integer, intent(in) :: a, b
            integer :: r
        end function pure_add
    end interface
end module pure_iface

submodule (pure_iface) pure_impl
    implicit none
contains
    module function pure_add(a, b) result(r)
        integer, intent(in) :: a, b
        integer :: r
        r = a + b
    end function pure_add
end submodule pure_impl

program t
    use pure_iface
    print *, pure_add(3, 4)
end program t
"#,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn submodule_intent_inout_argument() {
    let out = run_prints(
        r#"
module io_iface
    implicit none
    interface
        module subroutine double_it(x)
            integer, intent(inout) :: x
        end subroutine double_it
    end interface
end module io_iface

submodule (io_iface) io_impl
contains
    module subroutine double_it(x)
        integer, intent(inout) :: x
        x = x * 2
    end subroutine double_it
end submodule io_impl

program t
    use io_iface
    integer :: n = 6
    call double_it(n)
    print *, n
end program t
"#,
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn submodule_optional_and_out_arguments_fill_defaults() {
    let out = run_prints(
        r#"
module out_iface
    implicit none
    interface
        module subroutine fill_pair(base, a, b, result_out)
            integer, intent(in) :: base
            integer, optional, intent(in) :: a, b
            integer, intent(out) :: result_out
        end subroutine fill_pair
    end interface
end module out_iface

submodule (out_iface) out_impl
contains
    module subroutine fill_pair(base, a, b, result_out)
        integer, intent(in) :: base
        integer, optional, intent(in) :: a, b
        integer, intent(out) :: result_out
        if (present(a) .and. present(b)) then
            result_out = base + a + b
        else
            result_out = base + 1
        end if
    end subroutine fill_pair
end submodule out_impl

program t
    use out_iface
    integer :: x
    call fill_pair(6, result_out=x)
    print *, x
    call fill_pair(6, a=2, b=3, result_out=x)
    print *, x
end program t
"#,
    );
    assert_eq!(out, vec!["7", "11"]);
}

#[test]
fn submodule_can_reference_parent_implementation_helper() {
    let out = run_prints(
        r#"
module base_iface
    implicit none
contains
    integer function offset_base()
        offset_base = 20
    end function offset_base
    interface
        module function apply_offset(x) result(r)
            integer, intent(in) :: x
            integer :: r
        end function apply_offset
    end interface
end module base_iface

submodule (base_iface) base_impl
contains
    module function apply_offset(x) result(r)
        integer, intent(in) :: x
        integer :: r
        r = x + offset_base()
    end function apply_offset
end submodule base_impl

program t
    use base_iface
    print *, apply_offset(3)
end program t
"#,
    );
    assert_eq!(out, vec!["23"]);
}

#[test]
fn submodule_nested_interface_chain_with_child_helper() {
    let out = run_prints(
        r#"
module chain_iface
    implicit none
    interface
        module function top(x) result(r)
            integer, intent(in) :: x
            integer :: r
        end function top
    end interface
end module chain_iface

submodule (chain_iface) chain_mid
    interface
        module function helper(x) result(r)
            integer, intent(in) :: x
            integer :: r
        end function helper
    end interface
end submodule chain_mid

submodule (chain_iface:chain_mid) chain_leaf
    contains
    module function top(x) result(r)
        integer, intent(in) :: x
        integer :: r
        r = helper(x) + 3
    end function top

    module function helper(x) result(r)
        integer, intent(in) :: x
        integer :: r
        r = x * 2
    end function helper
end submodule chain_leaf

program t
    use chain_iface
    print *, top(4)
end program t
"#,
    );
    assert_eq!(out, vec!["11"]);
}

#[test]
fn submodule_allocatable_local_workspace() {
    let out = run_prints(
        r#"
module grow_iface
    implicit none
    interface
        module function grown(n) result(r)
            integer, intent(in) :: n
            integer :: r
        end function grown
    end interface
end module grow_iface

submodule (grow_iface) grow_impl
contains
    module function grown(n) result(r)
        integer, intent(in) :: n
        integer, allocatable :: buf(:)
        integer :: r
        allocate(buf(n))
        buf = 1
        r = sum(buf)
        deallocate(buf)
    end function grown
end submodule grow_impl

program t
    use grow_iface
    print *, grown(5)
end program t
"#,
    );
    assert_eq!(out, vec!["5"]);
}
