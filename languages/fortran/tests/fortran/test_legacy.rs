use super::helpers::compile_ok;

// ── GOTO and statement labels ─────────────────────────────────

#[test]
fn goto_basic() {
    compile_ok(
        r#"
program test
    integer :: x = 0
    goto 10
    x = 999
10  continue
    print *, x
end program test
"#,
    );
}

#[test]
fn goto_forward() {
    compile_ok(
        r#"
program test
    goto 20
10  print *, 'skip'
    goto 30
20  print *, 'landed'
30  continue
end program test
"#,
    );
}

#[test]
fn goto_in_loop() {
    compile_ok(
        r#"
program test
    integer :: i, s
    i = 1
    s = 0
10  if (i > 5) goto 20
    s = s + i
    i = i + 1
    goto 10
20  print *, s
end program test
"#,
    );
}

#[test]
fn computed_goto() {
    compile_ok(
        r#"
program test
    integer :: n = 2
    go to (10, 20, 30), n
10  print *, 'one'; goto 99
20  print *, 'two'; goto 99
30  print *, 'three'
99  continue
end program test
"#,
    );
}

#[test]
fn assigned_goto() {
    compile_ok(
        r#"
program test
    integer :: label
    assign 10 to label
    goto label
    print *, 'skipped'
10  continue
    print *, 'ok'
end program test
"#,
    );
}

// ── Arithmetic IF ─────────────────────────────────────────────

#[test]
fn arithmetic_if_negative() {
    compile_ok(
        r#"
program test
    real :: x = -1.0
    if (x) 10, 20, 30
10  print *, 'negative'; goto 99
20  print *, 'zero'; goto 99
30  print *, 'positive'
99  continue
end program test
"#,
    );
}

#[test]
fn arithmetic_if_zero() {
    compile_ok(
        r#"
program test
    real :: x = 0.0
    if (x) 10, 20, 30
10  print *, 'negative'; goto 99
20  print *, 'zero'; goto 99
30  print *, 'positive'
99  continue
end program test
"#,
    );
}

#[test]
fn arithmetic_if_positive() {
    compile_ok(
        r#"
program test
    real :: x = 1.0
    if (x) 10, 20, 30
10  print *, 'negative'; goto 99
20  print *, 'zero'; goto 99
30  print *, 'positive'
99  continue
end program test
"#,
    );
}

// ── Labeled DO loops ──────────────────────────────────────────

#[test]
fn labeled_do_basic() {
    compile_ok(
        r#"
program test
    integer :: i, s
    s = 0
    do 100 i = 1, 5
        s = s + i
100 continue
    print *, s
end program test
"#,
    );
}

#[test]
fn labeled_do_nested() {
    compile_ok(
        r#"
program test
    integer :: i, j, s
    s = 0
    do 200 i = 1, 3
        do 100 j = 1, 3
            s = s + 1
100     continue
200 continue
    print *, s
end program test
"#,
    );
}

#[test]
fn labeled_do_with_step() {
    compile_ok(
        r#"
program test
    integer :: i, s
    s = 0
    do 10 i = 0, 10, 2
        s = s + i
10  continue
    print *, s
end program test
"#,
    );
}

// ── CONTINUE ─────────────────────────────────────────────────

#[test]
fn continue_label() {
    compile_ok(
        r#"
program test
    integer :: i
    do i = 1, 5
        if (mod(i, 2) == 0) goto 100
        print *, i
100     continue
    end do
end program test
"#,
    );
}

// ── COMMON blocks ─────────────────────────────────────────────

#[test]
fn common_basic() {
    compile_ok(
        r#"
program test
    integer :: x, y
    common /data/ x, y
    x = 10
    y = 20
    print *, x + y
end program test
"#,
    );
}

#[test]
fn common_blank() {
    compile_ok(
        r#"
program test
    integer :: a, b
    common a, b
    a = 1
    b = 2
    print *, a * b
end program test
"#,
    );
}

#[test]
fn common_shared_subprogram() {
    compile_ok(
        r#"
program test
    integer :: total
    common /result/ total
    total = 0
    call accumulate(5)
    print *, total
contains
    subroutine accumulate(n)
        integer, intent(in) :: n
        integer :: total
        common /result/ total
        total = total + n
    end subroutine accumulate
end program test
"#,
    );
}

// ── EQUIVALENCE ───────────────────────────────────────────────

#[test]
fn equivalence_basic() {
    compile_ok(
        r#"
program test
    integer :: a
    integer :: b
    equivalence (a, b)
    a = 42
    print *, b
end program test
"#,
    );
}

#[test]
fn equivalence_array_scalar() {
    compile_ok(
        r#"
program test
    integer :: arr(4)
    integer :: first
    equivalence (arr(1), first)
    arr(1) = 99
    print *, first
end program test
"#,
    );
}

// ── DATA statements ───────────────────────────────────────────

#[test]
fn data_integer() {
    compile_ok(
        r#"
program test
    integer :: x, y
    data x /42/, y /99/
    print *, x + y
end program test
"#,
    );
}

#[test]
fn data_array() {
    compile_ok(
        r#"
program test
    integer :: a(5)
    data a /1, 2, 3, 4, 5/
    print *, a(3)
end program test
"#,
    );
}

#[test]
fn data_repeated() {
    compile_ok(
        r#"
program test
    integer :: a(6)
    data a /6*0/
    print *, a(1)
end program test
"#,
    );
}

#[test]
fn data_implied_do() {
    compile_ok(
        r#"
program test
    integer :: a(5)
    data (a(i), i=1,5) /1, 2, 3, 4, 5/
    print *, a(3)
end program test
"#,
    );
}

#[test]
fn data_character() {
    compile_ok(
        r#"
program test
    character(len=5) :: s
    data s /'hello'/
    print *, s
end program test
"#,
    );
}

#[test]
fn data_logical() {
    compile_ok(
        r#"
program test
    logical :: flag
    data flag /.true./
    print *, flag
end program test
"#,
    );
}

// ── BLOCK DATA ────────────────────────────────────────────────

#[test]
fn block_data_basic() {
    compile_ok(
        r#"
block data init_data
    integer :: x, y
    common /shared/ x, y
    data x /10/, y /20/
end block data init_data

program test
    integer :: x, y
    common /shared/ x, y
    print *, x + y
end program test
"#,
    );
}

// ── SAVE attribute ────────────────────────────────────────────

#[test]
fn save_basic() {
    compile_ok(
        r#"
program test
    call inc()
    call inc()
    call inc()
contains
    subroutine inc()
        integer, save :: count = 0
        count = count + 1
        print *, count
    end subroutine inc
end program test
"#,
    );
}

#[test]
fn save_array() {
    compile_ok(
        r#"
program test
    call store(42)
    call retrieve()
contains
    subroutine store(val)
        integer, intent(in) :: val
        integer, save :: stored
        stored = val
    end subroutine store
    subroutine retrieve()
        integer, save :: stored
        print *, stored
    end subroutine retrieve
end program test
"#,
    );
}

// ── EXTERNAL and INTRINSIC declarations ───────────────────────

#[test]
fn intrinsic_decl() {
    compile_ok(
        r#"
program test
    intrinsic :: sin, cos, sqrt
    real :: x
    x = sin(0.0)
    print *, x
end program test
"#,
    );
}

#[test]
fn external_decl() {
    compile_ok(
        r#"
program test
    external :: my_func
    print *, "ok"
contains
    function my_func(x)
        real :: my_func, x
        my_func = x * 2.0
    end function my_func
end program test
"#,
    );
}

// ── ENTRY statement ───────────────────────────────────────────

#[test]
fn entry_basic() {
    compile_ok(
        r#"
program test
    call init_and_run()
contains
    subroutine init_and_run()
        print *, 'init'
        return
    entry run_only()
        print *, 'run'
    end subroutine init_and_run
end program test
"#,
    );
}

// ── STOP with code ────────────────────────────────────────────

#[test]
fn stop_int() {
    compile_ok("program t\n  print *, 'before'\n  stop 0\nend program t\n");
}
#[test]
fn stop_string() {
    compile_ok("program t\n  stop 'clean exit'\nend program t\n");
}
#[test]
fn error_stop() {
    compile_ok(
        "program t\n  logical :: ok = .true.\n  if (.not. ok) error stop 'fatal'\n  print *, 'fine'\nend program t\n",
    );
}
