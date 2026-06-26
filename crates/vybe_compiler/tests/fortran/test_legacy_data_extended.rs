//! Extended legacy Fortran coverage: DATA implied-do, EQUIVALENCE overlap,
//! multi-variable COMMON, SAVE, and ENTRY. Distinct from `test_legacy.rs`
//! (basic DATA, EQUIVALENCE, COMMON, BLOCK DATA, SAVE, ENTRY).

use super::helpers::compile_ok;

fortran_cases! {
    // ── COMMON: multi-variable blank and named blocks ────────────────

    common_blank_four_sum => {
        "program t\ninteger :: a, b, c, d\ncommon a, b, c, d\na = 1; b = 2; c = 3; d = 4\nprint *, a + b + c + d\nend program t\n",
        ["10"]
    };

    common_blank_five_product => {
        "program t\ninteger :: v(5)\ncommon v\nv(1) = 2; v(2) = 3; v(3) = 1; v(4) = 4; v(5) = 5\nprint *, v(1) * v(2) * v(3)\nend program t\n",
        ["6"]
    };

    common_named_six_total => {
        "program t\ninteger :: i1, i2, i3, i4, i5, i6\ncommon /pool/ i1, i2, i3, i4, i5, i6\ni1 = 1; i2 = 2; i3 = 3; i4 = 4; i5 = 5; i6 = 6\nprint *, i1 + i2 + i3 + i4 + i5 + i6\nend program t\n",
        ["21"]
    };

    common_two_named_blocks => {
        "program t\ninteger :: ix\nreal :: rx\ncommon /ints/ ix\ncommon /reals/ rx\nix = 7\nrx = 2.5\nprint *, ix\nprint *, rx\nend program t\n",
        ["7", "2.5"]
    };

    common_array_three_sum => {
        "program t\ninteger :: arr(3)\ncommon /nums/ arr\narr(1) = 4; arr(2) = 5; arr(3) = 6\nprint *, sum(arr)\nend program t\n",
        ["15"]
    };

    common_real_integer_pair => {
        "program t\ninteger :: n\nreal :: x\ncommon /mix/ n, x\nn = 8\nx = 1.5\nprint *, n\nprint *, x\nend program t\n",
        ["8", "1.5"]
    };

    common_pair_sum_in_main => {
        "program t\ninteger :: base, extra\ncommon /pair/ base, extra\nbase = 6; extra = 4\nprint *, base + extra\nend program t\n",
        ["10"]
    };

    common_triple_assign_print => {
        "program t\ninteger :: p, q, r\ncommon /trio/ p, q, r\np = 11; q = 22; r = 33\nprint *, p\nprint *, q\nprint *, r\nend program t\n",
        ["11", "22", "33"]
    };

    common_append_second_declaration => {
        "program t\ninteger :: a, b, c\ncommon /grp/ a, b\ncommon /grp/ c\na = 1; b = 2; c = 3\nprint *, a + b + c\nend program t\n",
        ["6"]
    };

    common_seven_var_sum => {
        "program t\ninteger :: s(7)\ncommon /wide/ s\ns = [(i, i = 1, 7)]\nprint *, sum(s)\nend program t\n",
        ["28"]
    };

    common_real_triple_sum => {
        "program t\nreal :: u, v, w\ncommon /rv/ u, v, w\nu = 1.0; v = 2.0; w = 3.0\nprint *, u + v + w\nend program t\n",
        ["6"]
    };

    common_2d_array_total => {
        "program t\ninteger :: grid(2, 3)\ncommon /grid/ grid\ngrid = 1\nprint *, sum(grid)\nend program t\n",
        ["6"]
    };

    common_logical_pair_and => {
        "program t\nlogical :: f1, f2\ncommon /flags/ f1, f2\nf1 = .true.\nf2 = .false.\nprint *, f1 .and. f2\nprint *, f1 .or. f2\nend program t\n",
        ["false", "true"]
    };

    common_character_pair_len => {
        "program t\ncharacter(len=3) :: s1, s2\ncommon /txt/ s1, s2\ns1 = 'ab'\ns2 = 'cd'\nprint *, len_trim(s1)\nprint *, len_trim(s2)\nend program t\n",
        ["2", "2"]
    };

    common_four_int_product => {
        "program t\ninteger :: a, b, c, d\ncommon /quad/ a, b, c, d\na = 2; b = 3; c = 4; d = 5\nprint *, a * b * c * d\nend program t\n",
        ["120"]
    };

    common_blank_and_named => {
        "program t\ninteger :: a, b\nreal :: r\ncommon a, b\ncommon /one/ r\na = 3; b = 4; r = 0.5\nprint *, a + b\nprint *, r\nend program t\n",
        ["7", "0.5"]
    };

    // ── SAVE: initial value readable in same invocation ────────────

    save_initial_value_single_call => {
        "program t\ncall show()\ncontains\nsubroutine show()\ninteger, save :: tally = 17\nprint *, tally\nend subroutine show\nend program t\n",
        ["17"]
    };

    save_uninitialized_reads_zero => {
        "program t\ncall read_zero()\ncontains\nsubroutine read_zero()\ninteger, save :: bucket\nprint *, bucket\nend subroutine read_zero\nend program t\n",
        ["0"]
    };

    save_real_constant => {
        "program t\ncall pi_once()\ncontains\nsubroutine pi_once()\nreal, save :: pi = 3.25\nprint *, pi\nend subroutine pi_once\nend program t\n",
        ["3.25"]
    };

    save_logical_true_flag => {
        "program t\ncall flag_once()\ncontains\nsubroutine flag_once()\nlogical, save :: on = .true.\nprint *, on\nend subroutine flag_once\nend program t\n",
        ["true"]
    };

    save_array_three_elements => {
        "program t\ncall vec_once()\ncontains\nsubroutine vec_once()\ninteger, save :: vec(3) = (/5, 6, 7/)\nprint *, vec(2)\nprint *, sum(vec)\nend subroutine vec_once\nend program t\n",
        ["6", "18"]
    };

    // ── ENTRY: host entry executes when called directly ──────────────

    entry_host_entry_prints => {
        "program t\ncall worker()\ncontains\nsubroutine worker()\nprint *, 100\nreturn\nentry alt_worker()\nprint *, 200\nend subroutine worker\nend program t\n",
        ["100"]
    };
}

// ── DATA: implied-do and repeat extensions (compile-only) ─────────

#[test]
fn data_implied_do_with_step() {
    compile_ok(
        r#"
program t
    integer :: a(5)
    data (a(i), i = 1, 5, 2) /10, 20, 30/
    print *, a(1), a(3), a(5)
end program t
"#,
    );
}

#[test]
fn data_implied_do_nested_2d() {
    compile_ok(
        r#"
program t
    integer :: m(2, 3)
    data m /1, 2, 3, 4, 5, 6/
    print *, m(2, 2)
end program t
"#,
    );
}

#[test]
fn data_implied_do_two_variables() {
    compile_ok(
        r#"
program t
    integer :: x, y, z
    data x /1/, y /2/, z /3/
    print *, x + y + z
end program t
"#,
    );
}

#[test]
fn data_implied_do_with_repeat_factor() {
    compile_ok(
        r#"
program t
    integer :: v(6)
    data (v(i), i = 1, 6) /3*1, 3*9/
    print *, v(1), v(4)
end program t
"#,
    );
}

#[test]
fn data_implied_do_real_array() {
    compile_ok(
        r#"
program t
    real :: r(4)
    data (r(i), i = 1, 4) /1.5, 2.5, 3.5, 4.5/
    print *, r(2)
end program t
"#,
    );
}

#[test]
fn data_implied_do_character_array() {
    compile_ok(
        r#"
program t
    character(len=4) :: tags(2)
    data (tags(i), i = 1, 2) /'ab  ', 'cd  '/
    print *, tags(1)
end program t
"#,
    );
}

#[test]
fn data_implied_do_in_subroutine() {
    compile_ok(
        r#"
program t
    call init()
contains
    subroutine init()
        integer :: buf(3)
        data (buf(i), i = 1, 3) /4, 5, 6/
        print *, buf(2)
    end subroutine init
end program t
"#,
    );
}

#[test]
fn data_implied_do_multiple_sets() {
    compile_ok(
        r#"
program t
    integer :: a(3), b(2)
    data (a(i), i = 1, 3) /1, 2, 3/, (b(j), j = 1, 2) /8, 9/
    print *, a(1) + b(2)
end program t
"#,
    );
}

#[test]
fn data_implied_do_descending_step() {
    compile_ok(
        r#"
program t
    integer :: a(5)
    data (a(i), i = 5, 1, -1) /50, 40, 30, 20, 10/
    print *, a(5), a(1)
end program t
"#,
    );
}

#[test]
fn data_implied_do_partial_matrix_row() {
    compile_ok(
        r#"
program t
    integer :: table(3, 2)
    data (table(1, j), j = 1, 2) /7, 8/
    print *, table(1, 1)
end program t
"#,
    );
}

#[test]
fn data_implied_do_scalar_and_array_mix() {
    compile_ok(
        r#"
program t
    integer :: head, tail(2)
    data head /0/, (tail(i), i = 1, 2) /4, 5/
    print *, head + tail(2)
end program t
"#,
    );
}

// ── EQUIVALENCE: overlap and alias reads (compile-only) ─────────────

#[test]
fn equiv_read_after_write_overlap() {
    compile_ok(
        r#"
program t
    integer :: alpha, beta
    equivalence (alpha, beta)
    alpha = 77
    print *, beta
end program t
"#,
    );
}

#[test]
fn equiv_chain_three_names() {
    compile_ok(
        r#"
program t
    integer :: p, q, r
    equivalence (p, q, r)
    p = 12
    print *, q, r
end program t
"#,
    );
}

#[test]
fn equiv_adjacent_array_elements() {
    compile_ok(
        r#"
program t
    integer :: seq(4)
    equivalence (seq(2), seq(3))
    seq(2) = 55
    print *, seq(3)
end program t
"#,
    );
}

#[test]
fn equiv_scalar_to_array_element() {
    compile_ok(
        r#"
program t
    integer :: arr(3), scalar
    equivalence (arr(2), scalar)
    scalar = 31
    print *, arr(2)
end program t
"#,
    );
}

#[test]
fn equiv_multiple_independent_groups() {
    compile_ok(
        r#"
program t
    integer :: a, b, c, d
    equivalence (a, b)
    equivalence (c, d)
    a = 1
    c = 9
    print *, b, d
end program t
"#,
    );
}

#[test]
fn equiv_real_integer_overlap_read() {
    compile_ok(
        r#"
program t
    real :: x
    integer :: k
    equivalence (x, k)
    x = 2.0
    print *, k
end program t
"#,
    );
}

// ── COMMON: shared linkage and block-data shapes (compile-only) ─────

#[test]
fn common_shared_accumulator_subprogram() {
    compile_ok(
        r#"
program t
    integer :: total
    common /acc/ total
    total = 0
    call bump(4)
    call bump(6)
    print *, total
contains
    subroutine bump(n)
        integer, intent(in) :: n
        integer :: total
        common /acc/ total
        total = total + n
    end subroutine bump
end program t
"#,
    );
}

#[test]
fn common_function_and_program_slots() {
    compile_ok(
        r#"
program t
    integer :: n
    real :: rate
    common /cfg/ n, rate
    n = 5
    rate = 2.0
    print *, scaled()
contains
    function scaled()
        real :: scaled
        integer :: n
        real :: rate
        common /cfg/ n, rate
        scaled = n * rate
    end function scaled
end program t
"#,
    );
}

#[test]
fn block_data_two_common_blocks() {
    compile_ok(
        r#"
block data setup
    integer :: ix
    real :: rx
    common /ints/ ix
    common /reals/ rx
    data ix /42/, rx /3.5/
end block data setup

program t
    integer :: ix
    real :: rx
    common /ints/ ix
    common /reals/ rx
    print *, ix
    print *, rx
end program t
"#,
    );
}

// ── SAVE: cross-call persistence (compile-only) ─────────────────────

#[test]
fn save_persists_across_calls() {
    compile_ok(
        r#"
program t
    call tick()
    call tick()
    call tick()
contains
    subroutine tick()
        integer, save :: n = 0
        n = n + 1
        print *, n
    end subroutine tick
end program t
"#,
    );
}

#[test]
fn save_all_locals_listed() {
    compile_ok(
        r#"
program t
    call stash(3)
    call recall()
contains
    subroutine stash(v)
        integer, intent(in) :: v
        integer, save :: a, b
        a = v
        b = v * 2
    end subroutine stash
    subroutine recall()
        integer, save :: a, b
        print *, a, b
    end subroutine recall
end program t
"#,
    );
}

#[test]
fn save_in_function_result() {
    compile_ok(
        r#"
program t
    print *, counter()
    print *, counter()
contains
    function counter()
        integer, save :: n = 0
        integer :: counter
        n = n + 1
        counter = n
    end function counter
end program t
"#,
    );
}

// ── ENTRY: alternate entry points (compile-only) ────────────────────

#[test]
fn entry_with_dummy_arguments() {
    compile_ok(
        r#"
program t
    call master(5)
contains
    subroutine master(x)
        integer, intent(in) :: x
        print *, x
        return
    entry slave(y)
        integer :: y
        print *, y + 1
    end subroutine master
end program t
"#,
    );
}

#[test]
fn entry_call_alternate_name() {
    compile_ok(
        r#"
program t
    call door()
contains
    subroutine door()
        print *, 'main'
        return
    entry window()
        print *, 'alt'
    end subroutine door
end program t
"#,
    );
}

#[test]
fn entry_multiple_in_function() {
    compile_ok(
        r#"
program t
    call host(3)
contains
    subroutine host(n)
        integer, intent(in) :: n
        print *, n
        return
    entry host_alt(n)
        integer :: n
        print *, n + 10
    end subroutine host
end program t
"#,
    );
}

#[test]
fn entry_nested_return_paths() {
    compile_ok(
        r#"
program t
    call pipeline()
contains
    subroutine pipeline()
        print *, 1
        return
    entry pipeline_b()
        print *, 2
        return
    entry pipeline_c()
        print *, 3
    end subroutine pipeline
end program t
"#,
    );
}

#[test]
fn entry_function_primary_body() {
    compile_ok(
        r#"
program t
    print *, primary(4)
contains
    function primary(n)
        integer, intent(in) :: n
        integer :: primary
        primary = n * 2
    entry backup(n)
        integer :: n
        primary = n + 1
    end function primary
end program t
"#,
    );
}
