//! Extended module USE coverage: only lists, rename, public/private access,
//! module variables and procedures from program units, and interface blocks.

use super::helpers::compile_ok;

fortran_cases! {
    // ── USE, ONLY ────────────────────────────────────────────────────

    use_only_function_from_module => {
        "module calc\nimplicit none\ncontains\nfunction triple(n) result(r)\ninteger, intent(in) :: n\ninteger :: r\nr = n * 3\nend function triple\nend module calc\nprogram t\nuse calc, only: triple\nprint *, triple(4)\nend program t\n",
        ["12"]
    };

    use_only_subroutine_from_module => {
        "module io_mod\nimplicit none\ncontains\nsubroutine emit(v)\ninteger, intent(in) :: v\nprint *, v\nend subroutine emit\nend module io_mod\nprogram t\nuse io_mod, only: emit\ncall emit(77)\nend program t\n",
        ["77"]
    };

    use_only_real_module_variable => {
        "module phys\nimplicit none\nreal :: gravity = 9.8\nreal :: mass = 2.0\nend module phys\nprogram t\nuse phys, only: gravity\nprint *, int(gravity)\nend program t\n",
        ["9"]
    };

    use_only_logical_flag => {
        "module flags\nimplicit none\nlogical :: active = .true.\nlogical :: debug = .false.\nend module flags\nprogram t\nuse flags, only: active\nprint *, active\nend program t\n",
        ["1"]
    };

    use_only_character_name => {
        "module names\nimplicit none\ncharacter(len=6) :: tag = 'alpha'\ncharacter(len=6) :: alt = 'beta'\nend module names\nprogram t\nuse names, only: tag\nprint *, trim(tag)\nend program t\n",
        ["alpha"]
    };

    use_only_parameter_constant => {
        "module limits\nimplicit none\ninteger, parameter :: MAX_N = 50\ninteger, parameter :: MIN_N = 1\nend module limits\nprogram t\nuse limits, only: MAX_N\nprint *, MAX_N\nend program t\n",
        ["50"]
    };

    use_only_one_procedure_of_two => {
        "module arith\nimplicit none\ncontains\nfunction add2(a, b) result(r)\ninteger, intent(in) :: a, b\ninteger :: r\nr = a + b\nend function add2\nfunction sub2(a, b) result(r)\ninteger, intent(in) :: a, b\ninteger :: r\nr = a - b\nend function sub2\nend module arith\nprogram t\nuse arith, only: add2\nprint *, add2(9, 4)\nend program t\n",
        ["13"]
    };

    use_only_type_symbol => {
        "module geom\nimplicit none\ntype :: Point\nreal :: x, y\nend type Point\ncontains\nfunction origin() result(p)\ntype(Point) :: p\np%x = 0.0\np%y = 0.0\nend function origin\nend module geom\nprogram t\nuse geom, only: Point\ntype(Point) :: p\np%x = 3.0\np%y = 4.0\nprint *, int(p%x)\nend program t\n",
        ["3"]
    };

    // ── USE rename ───────────────────────────────────────────────────

    rename_function_to_short_alias => {
        "module mathfn\nimplicit none\ncontains\nfunction cube(n) result(r)\ninteger, intent(in) :: n\ninteger :: r\nr = n * n * n\nend function cube\nend module mathfn\nprogram t\nuse mathfn, cb => cube\nprint *, cb(3)\nend program t\n",
        ["27"]
    };

    rename_subroutine_to_local_name => {
        "module outmod\nimplicit none\ncontains\nsubroutine show_value(v)\ninteger, intent(in) :: v\nprint *, v\nend subroutine show_value\nend module outmod\nprogram t\nuse outmod, disp => show_value\ncall disp(19)\nend program t\n",
        ["19"]
    };

    rename_integer_variable_alias => {
        "module counters\nimplicit none\ninteger :: tally = 8\ninteger :: spare = 0\nend module counters\nprogram t\nuse counters, total => tally\nprint *, total\nend program t\n",
        ["8"]
    };

    rename_parameter_alias => {
        "module units\nimplicit none\nreal, parameter :: METERS_PER_MILE = 1609.34\nend module units\nprogram t\nuse units, mile_m => METERS_PER_MILE\nprint *, int(mile_m)\nend program t\n",
        ["1609"]
    };

    rename_combined_with_only_clause => {
        "module pack\nimplicit none\ninteger :: hidden = 1\ninteger :: visible = 2\ncontains\nfunction pick(x) result(r)\ninteger, intent(in) :: x\ninteger :: r\nr = x + hidden\nend function pick\nend module pack\nprogram t\nuse pack, only: choose => pick\nprint *, choose(5)\nend program t\n",
        ["6"]
    };

    rename_two_symbols_in_one_use => {
        "module pair\nimplicit none\ninteger :: first = 10\ninteger :: second = 3\ncontains\nfunction sum_pair() result(r)\ninteger :: r\nr = first + second\nend function sum_pair\nend module pair\nprogram t\nuse pair, a => first, b => second, total => sum_pair\nprint *, total()\nend program t\n",
        ["13"]
    };

    // ── PUBLIC / PRIVATE ─────────────────────────────────────────────

    public_variable_read_from_program => {
        "module settings\nimplicit none\ninteger, public :: level = 5\nend module settings\nprogram t\nuse settings\nprint *, level\nend program t\n",
        ["5"]
    };

    private_var_exposed_via_getter => {
        "module vault\nimplicit none\nprivate\npublic :: read_vault\ninteger :: stored = 99\ncontains\nfunction read_vault() result(v)\ninteger :: v\nv = stored\nend function read_vault\nend module vault\nprogram t\nuse vault\nprint *, read_vault()\nend program t\n",
        ["99"]
    };

    default_public_with_one_private_symbol => {
        "module mix\nimplicit none\ninteger, public :: open_val = 7\ninteger, private :: closed_val = 100\ncontains\nfunction expose_open() result(v)\ninteger :: v\nv = open_val\nend function expose_open\nend module mix\nprogram t\nuse mix\nprint *, expose_open()\nend program t\n",
        ["7"]
    };

    explicit_public_subroutine_callable => {
        "module greet\nimplicit none\ncontains\npublic :: say_hi\nsubroutine say_hi()\nprint *, 1\nend subroutine say_hi\nend module greet\nprogram t\nuse greet\ncall say_hi()\nend program t\n",
        ["1"]
    };

    public_function_returns_product => {
        "module prodmod\nimplicit none\ncontains\npublic :: product2\nfunction product2(a, b) result(r)\ninteger, intent(in) :: a, b\ninteger :: r\nr = a * b\nend function product2\nend module prodmod\nprogram t\nuse prodmod\nprint *, product2(6, 7)\nend program t\n",
        ["42"]
    };

    private_procedure_public_facade => {
        "module facade\nimplicit none\nprivate\npublic :: facade_double\ncontains\nfunction inner_double(n) result(r)\ninteger, intent(in) :: n\ninteger :: r\nr = n * 2\nend function inner_double\nfunction facade_double(n) result(r)\ninteger, intent(in) :: n\ninteger :: r\nr = inner_double(n)\nend function facade_double\nend module facade\nprogram t\nuse facade\nprint *, facade_double(11)\nend program t\n",
        ["22"]
    };

    public_constants_both_visible => {
        "module constpair\nimplicit none\ninteger, public, parameter :: LOW = 2\ninteger, public, parameter :: HIGH = 9\nend module constpair\nprogram t\nuse constpair\nprint *, LOW + HIGH\nend program t\n",
        ["11"]
    };

    // ── Module procedures called from program ────────────────────────

    program_calls_module_subroutine => {
        "module actions\nimplicit none\ncontains\nsubroutine bump(n)\ninteger, intent(inout) :: n\nn = n + 1\nend subroutine bump\nend module actions\nprogram t\nuse actions\ninteger :: x = 4\ncall bump(x)\nprint *, x\nend program t\n",
        ["5"]
    };

    program_calls_module_function => {
        "module funcs\nimplicit none\ncontains\nfunction halve(n) result(r)\nreal, intent(in) :: n\nreal :: r\nr = n / 2.0\nend function halve\nend module funcs\nprogram t\nuse funcs\nprint *, int(halve(14.0))\nend program t\n",
        ["7"]
    };

    program_calls_module_logical_function => {
        "module checks\nimplicit none\ncontains\nfunction is_zero(n) result(b)\ninteger, intent(in) :: n\nlogical :: b\nb = n == 0\nend function is_zero\nend module checks\nprogram t\nuse checks\nprint *, is_zero(0)\nend program t\n",
        ["1"]
    };

    program_calls_module_character_function => {
        "module labels\nimplicit none\ncontains\nfunction label_for(n) result(s)\ninteger, intent(in) :: n\ncharacter(len=4) :: s\nif (n == 1) then\ns = 'one'\nelse\ns = 'many'\nend if\nend function label_for\nend module labels\nprogram t\nuse labels\nprint *, trim(label_for(1))\nend program t\n",
        ["one"]
    };

    program_calls_two_module_procedures => {
        "module duo\nimplicit none\ncontains\nfunction inc(n) result(r)\ninteger, intent(in) :: n\ninteger :: r\nr = n + 1\nend function inc\nfunction dec(n) result(r)\ninteger, intent(in) :: n\ninteger :: r\nr = n - 1\nend function dec\nend module duo\nprogram t\nuse duo\nprint *, inc(dec(6))\nend program t\n",
        ["6"]
    };

    module_subroutine_with_no_arguments => {
        "module ping\nimplicit none\ncontains\nsubroutine ping_once()\nprint *, 0\nend subroutine ping_once\nend module ping\nprogram t\nuse ping\ncall ping_once()\nend program t\n",
        ["0"]
    };

    module_procedure_chain_three_calls => {
        "module chain\nimplicit none\ncontains\nfunction step1(n) result(r)\ninteger, intent(in) :: n\ninteger :: r\nr = n + 1\nend function step1\nfunction step2(n) result(r)\ninteger, intent(in) :: n\ninteger :: r\nr = step1(n) + 1\nend function step2\nend module chain\nprogram t\nuse chain\nprint *, step2(3)\nend program t\n",
        ["5"]
    };

    module_subroutine_intent_out_fill => {
        "module fill\nimplicit none\ncontains\nsubroutine fill_pair(a, b)\ninteger, intent(out) :: a, b\na = 2\nb = 5\nend subroutine fill_pair\nend module fill\nprogram t\nuse fill\ninteger :: x, y\ncall fill_pair(x, y)\nprint *, x + y\nend program t\n",
        ["7"]
    };

    // ── Module variables ─────────────────────────────────────────────

    module_integer_variable_in_program => {
        "module ints\nimplicit none\ninteger :: seed = 13\nend module ints\nprogram t\nuse ints\nprint *, seed\nend program t\n",
        ["13"]
    };

    module_real_variable_in_expression => {
        "module reals\nimplicit none\nreal :: rate = 2.5\nend module reals\nprogram t\nuse reals\nprint *, int(rate * 4.0)\nend program t\n",
        ["10"]
    };

    module_array_public_element => {
        "module arrmod\nimplicit none\ninteger :: data(3) = [4, 5, 6]\nend module arrmod\nprogram t\nuse arrmod\nprint *, data(2)\nend program t\n",
        ["5"]
    };

    module_parameters_in_area_formula => {
        "module shapes\nimplicit none\nreal, parameter :: PI = 3.0\nreal, parameter :: R = 2.0\nend module shapes\nprogram t\nuse shapes\nprint *, int(PI * R * R)\nend program t\n",
        ["12"]
    };

    module_variable_updated_by_procedure => {
        "module tallymod\nimplicit none\ninteger :: hits = 0\ncontains\nsubroutine register_hit()\nhits = hits + 1\nend subroutine register_hit\nend module tallymod\nprogram t\nuse tallymod\ncall register_hit()\ncall register_hit()\nprint *, hits\nend program t\n",
        ["2"]
    };

    module_two_public_vars_summed => {
        "module pairvals\nimplicit none\ninteger :: left = 6\ninteger :: right = 8\nend module pairvals\nprogram t\nuse pairvals\nprint *, left + right\nend program t\n",
        ["14"]
    };

    // ── Interface blocks ─────────────────────────────────────────────

    module_generic_minimum_integers => {
        "module minmod\nimplicit none\ninterface my_min\nmodule procedure min_int\nend interface\ncontains\nfunction min_int(a, b) result(r)\ninteger, intent(in) :: a, b\ninteger :: r\nif (a < b) then\nr = a\nelse\nr = b\nend if\nend function min_int\nend module minmod\nprogram t\nuse minmod\nprint *, my_min(8, 3)\nend program t\n",
        ["3"]
    };

    module_generic_max_mixed_kinds => {
        "module maxmod\nimplicit none\ninterface my_max\nmodule procedure max_int, max_real\nend interface\ncontains\nfunction max_int(a, b) result(r)\ninteger, intent(in) :: a, b\ninteger :: r\nr = max(a, b)\nend function max_int\nfunction max_real(a, b) result(r)\nreal, intent(in) :: a, b\nreal :: r\nr = max(a, b)\nend function max_real\nend module maxmod\nprogram t\nuse maxmod\nprint *, my_max(2, 9)\nprint *, int(my_max(1.5, 3.7))\nend program t\n",
        ["9", "3"]
    };

    module_interface_operator_unary_negate => {
        "module neg\nimplicit none\ntype :: Signed\ninteger :: v\nend type Signed\ninterface operator(-)\nmodule procedure negate_signed\nend interface\ncontains\nfunction negate_signed(a) result(b)\ntype(Signed), intent(in) :: a\ntype(Signed) :: b\nb%v = -a%v\nend function negate_signed\nend module neg\nprogram t\nuse neg\ntype(Signed) :: x, y\nx%v = 7\ny = -x\nprint *, y%v\nend program t\n",
        ["-7"]
    };

    module_interface_resolves_external_shape => {
        "module iface_use\nimplicit none\ninterface\nfunction extern_double(n) result(r)\ninteger, intent(in) :: n\ninteger :: r\nend function extern_double\nend interface\ncontains\nfunction call_double(n) result(r)\ninteger, intent(in) :: n\ninteger :: r\nr = extern_double(n)\nend function call_double\nend module iface_use\nfunction extern_double(n) result(r)\ninteger, intent(in) :: n\ninteger :: r\nr = n * 2\nend function extern_double\nprogram t\nuse iface_use\nprint *, call_double(6)\nend program t\n",
        ["12"]
    };

    module_generic_len_overloads => {
        "module lenmod\nimplicit none\ninterface my_len\nmodule procedure len_int, len_char\nend interface\ncontains\nfunction len_int(n) result(r)\ninteger, intent(in) :: n\ninteger :: r\nr = 1\nend function len_int\nfunction len_char(s) result(r)\ncharacter(len=*), intent(in) :: s\ninteger :: r\nr = len_trim(s)\nend function len_char\nend module lenmod\nprogram t\nuse lenmod\nprint *, my_len(0)\nprint *, my_len('abcd')\nend program t\n",
        ["1", "4"]
    };
}

// ── Compile-only module USE / interface shapes ─────────────────────

#[test]
fn compile_use_only_triple_symbol_list() {
    compile_ok(
        r#"
module trio
    implicit none
    integer :: u = 1, v = 2, w = 3
contains
    function sum3() result(r)
        integer :: r
        r = u + v + w
    end function sum3
end module trio

program t
    use trio, only: u, w, sum3
    print *, sum3()
end program t
"#,
    );
}

#[test]
fn compile_rename_subroutine_with_only() {
    compile_ok(
        r#"
module runner
    implicit none
contains
    subroutine execute_job()
        print *, "done"
    end subroutine execute_job
end module runner

program t
    use runner, only: go => execute_job
    call go()
end program t
"#,
    );
}

#[test]
fn compile_private_default_many_public_entities() {
    compile_ok(
        r#"
module access_mix
    implicit none
    private
    public :: a, b, show
    integer :: a = 1
    integer :: b = 2
    integer :: c = 3
contains
    subroutine show()
        print *, a + b
    end subroutine show
end module access_mix

program t
    use access_mix
    call show()
end program t
"#,
    );
}

#[test]
fn compile_module_interface_procedure_signature() {
    compile_ok(
        r#"
module sig_iface
    implicit none
    interface
        module function signed_add(a, b) result(r)
            integer, intent(in) :: a, b
            integer :: r
        end function signed_add
    end interface
end module sig_iface

program t
    print *, "ok"
end program t
"#,
    );
}

#[test]
fn compile_interface_assignment_int_to_real() {
    compile_ok(
        r#"
module assign_iface
    implicit none
    interface assignment(=)
        module procedure int_assign_real
    end interface
contains
    subroutine int_assign_real(r, i)
        real, intent(out) :: r
        integer, intent(in) :: i
        r = real(i)
    end subroutine int_assign_real
end module assign_iface

program t
    use assign_iface
    real :: x
    x = 5
    print *, x
end program t
"#,
    );
}

#[test]
fn compile_interface_operator_equality() {
    compile_ok(
        r#"
module eq_iface
    implicit none
    type :: Tag
        integer :: id
    end type Tag
    interface operator(==)
        module procedure tags_equal
    end interface
contains
    function tags_equal(a, b) result(same)
        type(Tag), intent(in) :: a, b
        logical :: same
        same = a%id == b%id
    end function tags_equal
end module eq_iface

program t
    use eq_iface
    type(Tag) :: x, y
    x%id = 1
    y%id = 1
    print *, x == y
end program t
"#,
    );
}

#[test]
fn compile_module_many_contains_procedures() {
    compile_ok(
        r#"
module toolbox
    implicit none
contains
    function f1(x) result(r)
        integer, intent(in) :: x
        integer :: r
        r = x
    end function f1
    function f2(x) result(r)
        integer, intent(in) :: x
        integer :: r
        r = x + 1
    end function f2
    subroutine noop()
    end subroutine noop
end module toolbox

program t
    use toolbox
    print *, f1(1) + f2(2)
end program t
"#,
    );
}

#[test]
fn compile_use_rename_and_only_combined() {
    compile_ok(
        r#"
module symbols
    implicit none
    integer :: alpha = 1
    integer :: beta = 2
contains
    function combine() result(r)
        integer :: r
        r = alpha + beta
    end function combine
end module symbols

program t
    use symbols, only: total => combine
    print *, total()
end program t
"#,
    );
}

#[test]
fn compile_interface_external_subroutine_block() {
    compile_ok(
        r#"
module ext_iface
    implicit none
    interface
        subroutine external_log(msg)
            character(len=*), intent(in) :: msg
        end subroutine external_log
    end interface
contains
    subroutine relay(msg)
        character(len=*), intent(in) :: msg
        call external_log(msg)
    end subroutine relay
end module ext_iface

subroutine external_log(msg)
    character(len=*), intent(in) :: msg
    print *, trim(msg)
end subroutine external_log

program t
    use ext_iface
    call relay("hi")
end program t
"#,
    );
}

#[test]
fn compile_generic_interface_three_procedures() {
    compile_ok(
        r#"
module g3
    implicit none
    interface pick
        module procedure pick_int, pick_real, pick_logical
    end interface
contains
    function pick_int(v) result(r)
        integer, intent(in) :: v
        integer :: r
        r = v
    end function pick_int
    function pick_real(v) result(r)
        real, intent(in) :: v
        real :: r
        r = v
    end function pick_real
    function pick_logical(v) result(r)
        logical, intent(in) :: v
        logical :: r
        r = v
    end function pick_logical
end module g3

program t
    use g3
    print *, pick(1)
    print *, int(pick(2.0))
    print *, pick(.true.)
end program t
"#,
    );
}

#[test]
fn compile_public_private_mixed_module_entities() {
    compile_ok(
        r#"
module mixed_access
    implicit none
    private
    public :: visible_val, reveal
    integer :: hidden = 0
    integer, public :: visible_val = 4
contains
    function reveal() result(v)
        integer :: v
        v = hidden + visible_val
    end function reveal
end module mixed_access

program t
    use mixed_access
    print *, reveal()
end program t
"#,
    );
}

#[test]
fn compile_module_interface_block_standalone() {
    compile_ok(
        r#"
module standalone_iface
    implicit none
    interface
        function area_circle(r) result(a)
            real, intent(in) :: r
            real :: a
        end function area_circle
    end interface
contains
    function scaled_area(r, s) result(a)
        real, intent(in) :: r, s
        real :: a
        a = area_circle(r) * s
    end function scaled_area
end module standalone_iface

function area_circle(r) result(a)
    real, intent(in) :: r
    real :: a
    a = 3.0 * r * r
end function area_circle

program t
    use standalone_iface
    print *, int(scaled_area(2.0, 1.0))
end program t
"#,
    );
}
