use super::helpers::run_prints;

// ── DO WHILE basic ────────────────────────────────────────────

#[test]
fn do_while_basic() {
    let out = run_prints(
        r#"
program test
    integer :: n = 0
    do while (n < 5)
        n = n + 1
    end do
    print *, n
end program test
"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn do_while_count_down() {
    let out = run_prints(
        r#"
program test
    integer :: n = 10
    do while (n > 0)
        n = n - 1
    end do
    print *, n
end program test
"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn do_while_accumulate() {
    let out = run_prints(
        r#"
program test
    integer :: i = 1, s = 0
    do while (i <= 10)
        s = s + i
        i = i + 1
    end do
    print *, s
end program test
"#,
    );
    assert_eq!(out, vec!["55"]);
}

#[test]
fn do_while_inline_body() {
    let out = run_prints(
        r#"
program test
    integer :: i = 0
    do while (i < 3); i = i + 1; end do
    print *, i
end program test
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn do_while_condition_false_initially() {
    let out = run_prints(
        r#"
program test
    integer :: n = 10
    do while (n < 5)
        n = n + 1
    end do
    print *, n
end program test
"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn do_while_with_logical_var() {
    let out = run_prints(
        r#"
program test
    integer :: n = 0
    logical :: keep_going
    keep_going = .true.
    do while (keep_going)
        n = n + 1
        if (n >= 5) keep_going = .false.
    end do
    print *, n
end program test
"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn do_while_real_condition() {
    let out = run_prints(
        r#"
program test
    real :: x = 1.0
    integer :: count = 0
    do while (x < 1000.0)
        x = x * 2.0
        count = count + 1
    end do
    print *, count
end program test
"#,
    );
    assert_eq!(out, vec!["10"]);
}

// ── DO WHILE with EXIT ────────────────────────────────────────

#[test]
fn do_while_with_exit() {
    let out = run_prints(
        r#"
program test
    integer :: n = 0
    do while (.true.)
        n = n + 1
        if (n == 7) exit
    end do
    print *, n
end program test
"#,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn do_while_exit_on_condition() {
    let out = run_prints(
        r#"
program test
    integer :: i = 0, s = 0
    do while (i < 100)
        i = i + 1
        s = s + i
        if (s > 50) exit
    end do
    print *, s
end program test
"#,
    );
    assert_eq!(out, vec!["55"]);
}

// ── DO WHILE with CYCLE ───────────────────────────────────────

#[test]
fn do_while_with_cycle() {
    let out = run_prints(
        r#"
program test
    integer :: i = 0, s = 0
    do while (i < 10)
        i = i + 1
        if (mod(i, 2) == 0) cycle
        s = s + i
    end do
    print *, s
end program test
"#,
    );
    assert_eq!(out, vec!["25"]);
}

#[test]
fn do_while_cycle_skip_zero() {
    let out = run_prints(
        r#"
program test
    integer :: i = -3, count = 0
    do while (i <= 3)
        i = i + 1
        if (i == 0) cycle
        count = count + 1
    end do
    print *, count
end program test
"#,
    );
    assert_eq!(out, vec!["5"]);
}

// ── Nested DO WHILE ───────────────────────────────────────────

#[test]
fn nested_do_while() {
    let out = run_prints(
        r#"
program test
    integer :: i = 1, j, s
    s = 0
    do while (i <= 3)
        j = 1
        do while (j <= 3)
            s = s + 1
            j = j + 1
        end do
        i = i + 1
    end do
    print *, s
end program test
"#,
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn nested_do_while_with_exit() {
    let out = run_prints(
        r#"
program test
    integer :: i = 0, j, count
    count = 0
    do while (i < 5)
        i = i + 1
        j = 0
        do while (j < 5)
            j = j + 1
            if (j == 3) exit
            count = count + 1
        end do
    end do
    print *, count
end program test
"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn do_while_named_loop_and_inner_break_condition() {
    let out = run_prints(
        r#"
program test
    integer :: i, j, s
    s = 0
    i = 0
    outer: do while (i < 5)
        i = i + 1
        j = 0
        do while (j < 5)
            j = j + 1
            if (j == 3) exit
            s = s + 1
        end do
    end do outer
    print *, s
end program test
"#,
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn do_while_with_cycle_and_named_exit() {
    let out = run_prints(
        r#"
program test
    integer :: i, n
    n = 0
    i = 0
    limit: do while (i < 20)
        i = i + 1
        if (mod(i, 2) == 0) cycle
        n = n + 1
        if (n == 4) exit limit
    end do limit
    print *, i
    print *, n
end program test
"#,
    );
    assert_eq!(out, vec!["7", "4"]);
}

// ── DO WHILE with complex logical condition ───────────────────

#[test]
fn do_while_and_condition() {
    let out = run_prints(
        r#"
program test
    integer :: x = 0, y = 10
    do while (x < 5 .and. y > 5)
        x = x + 1
        y = y - 1
    end do
    print *, x
    print *, y
end program test
"#,
    );
    assert_eq!(out, vec!["5", "5"]);
}

#[test]
fn do_while_or_condition() {
    let out = run_prints(
        r#"
program test
    integer :: a = 0, b = 0
    do while (a < 3 .or. b < 3)
        if (a < 3) a = a + 1
        if (b < 3) b = b + 1
    end do
    print *, a
    print *, b
end program test
"#,
    );
    assert_eq!(out, vec!["3", "3"]);
}

#[test]
fn do_while_not_condition() {
    let out = run_prints(
        r#"
program test
    integer :: n = 0
    logical :: done
    done = .false.
    do while (.not. done)
        n = n + 1
    if (n == 5) done = .true.
    end do
    print *, n
end program test
"#,
    );
    assert_eq!(out, vec!["5"]);
}

// ── DO WHILE reading data ─────────────────────────────────────

#[test]
fn do_while_read_until_eof() {
    let out = run_prints(
        r#"
program test
    integer :: n, ios, s, u, i
    integer, dimension(4) :: nums
    nums = [1, 2, 3, 4]
    s = 0
    open(newunit=u, status='scratch', action='readwrite')
    do i = 1, 4
        write(u, '(I0)') nums(i)
    end do
    rewind(u)
    do while (.true.)
        read(u, *, iostat=ios) n
        if (ios /= 0) exit
        s = s + n
    end do
    print *, s
    close(u)
end program test
"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn do_while_with_logical_not_condition() {
    let out = run_prints(
        r#"
program test
    logical :: keep_running
    integer :: n
    keep_running = .true.
    n = 0
    do while (.not. .not. keep_running)
        n = n + 1
        if (n == 3) keep_running = .false.
        if (n > 5) exit
    end do
    print *, n
end program test
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn do_while_and_condition_with_parentheses() {
    let out = run_prints(
        r#"
program test
    integer :: a
    integer :: b
    a = 0
    b = 0
    do while (a < 5 .and. (b < 2))
        a = a + 1
        b = b + 1
    end do
    print *, a
    print *, b
end program test
        "#,
    );
    assert_eq!(out, vec!["2", "2"]);
}

// ── DO WHILE in subroutine ────────────────────────────────────

#[test]
fn do_while_in_subroutine() {
    let out = run_prints(
        r#"
program test
    integer :: result
    call compute(result)
    print *, result
contains
    subroutine compute(r)
        integer, intent(out) :: r
        integer :: n = 0
        r = 0
        do while (n < 10)
            n = n + 1
            r = r + n
        end do
    end subroutine compute
end program test
"#,
    );
    assert_eq!(out, vec!["55"]);
}

#[test]
fn do_while_factorial() {
    let out = run_prints(
        r#"
program test
    integer :: n = 5, f = 1
    do while (n > 1)
        f = f * n
        n = n - 1
    end do
    print *, f
end program test
"#,
    );
    assert_eq!(out, vec!["120"]);
}

#[test]
fn do_while_fibonacci() {
    let out = run_prints(
        r#"
program test
    integer :: a = 0, b = 1, tmp, count = 0
    do while (b < 100)
        tmp = a + b
        a = b
        b = tmp
        count = count + 1
    end do
    print *, b
end program test
"#,
    );
    assert_eq!(out, vec!["144"]);
}

#[test]
fn do_while_newton_sqrt() {
    let out = run_prints(
        r#"
program test
    real :: x = 2.0, g = 1.0, prev
    integer :: it
    it = 0
    do while (.true.)
        prev = g
        g = 0.5 * (g + x / g)
        it = it + 1
        if (abs(g - prev) < 1e-7 .or. it > 20) exit
    end do
    if (abs(g * g - x) < 1.0e-3) then
        print *, 1
    else
        print *, 0
    end if
end program test
"#,
    );
    assert_eq!(out, vec!["1"]);
}

// ── Named DO WHILE (covered in test_named_loops but exercised here) ──

#[test]
fn do_while_two_variables() {
    let out = run_prints(
        r#"
program test
    integer :: m = 1, n = 256
    integer :: steps = 0
    do while (m < n)
        m = m * 2
        steps = steps + 1
    end do
    print *, steps
end program test
"#,
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn do_while_named_outer_exit_from_inner() {
    let out = run_prints(
        r#"
program test
    integer :: outer, inner, total
    outer = 0
    total = 0
    pump: do while (outer < 10)
        outer = outer + 1
        inner = 0
        do while (inner < 5)
            inner = inner + 1
            if (outer == 2 .and. inner == 3) exit pump
            total = total + 1
        end do
    end do pump
    print *, outer
    print *, total
end program test
"#,
    );
    assert_eq!(out, vec!["2", "4"]);
}

#[test]
fn do_while_named_cycle_skips_selected_steps() {
    let out = run_prints(
        r#"
program test
    integer :: i, total
    i = 0
    total = 0
    spin: do while (i < 6)
        i = i + 1
        if (mod(i, 3) == 0) cycle spin
        total = total + 1
    end do spin
    print *, i
    print *, total
end program test
"#,
    );
    assert_eq!(out, vec!["6", "4"]);
}
