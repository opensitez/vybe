use super::helpers::compile_ok;

// ── DO WHILE basic ────────────────────────────────────────────

#[test] fn do_while_basic() {
    compile_ok(r#"
program test
    integer :: n = 0
    do while (n < 5)
        n = n + 1
    end do
    print *, n
end program test
"#);
}

#[test] fn do_while_count_down() {
    compile_ok(r#"
program test
    integer :: n = 10
    do while (n > 0)
        n = n - 1
    end do
    print *, n
end program test
"#);
}

#[test] fn do_while_accumulate() {
    compile_ok(r#"
program test
    integer :: i = 1, s = 0
    do while (i <= 10)
        s = s + i
        i = i + 1
    end do
    print *, s
end program test
"#);
}

#[test] fn do_while_condition_false_initially() {
    compile_ok(r#"
program test
    integer :: n = 10
    do while (n < 5)
        n = n + 1
    end do
    print *, n
end program test
"#);
}

#[test] fn do_while_with_logical_var() {
    compile_ok(r#"
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
"#);
}

#[test] fn do_while_real_condition() {
    compile_ok(r#"
program test
    real :: x = 1.0
    integer :: count = 0
    do while (x < 1000.0)
        x = x * 2.0
        count = count + 1
    end do
    print *, count
end program test
"#);
}

// ── DO WHILE with EXIT ────────────────────────────────────────

#[test] fn do_while_with_exit() {
    compile_ok(r#"
program test
    integer :: n = 0
    do while (.true.)
        n = n + 1
        if (n == 7) exit
    end do
    print *, n
end program test
"#);
}

#[test] fn do_while_exit_on_condition() {
    compile_ok(r#"
program test
    integer :: i = 0, s = 0
    do while (i < 100)
        i = i + 1
        s = s + i
        if (s > 50) exit
    end do
    print *, s > 50
end program test
"#);
}

// ── DO WHILE with CYCLE ───────────────────────────────────────

#[test] fn do_while_with_cycle() {
    compile_ok(r#"
program test
    integer :: i = 0, s = 0
    do while (i < 10)
        i = i + 1
        if (mod(i, 2) == 0) cycle
        s = s + i
    end do
    print *, s
end program test
"#);
}

#[test] fn do_while_cycle_skip_zero() {
    compile_ok(r#"
program test
    integer :: i = -3, count = 0
    do while (i <= 3)
        i = i + 1
        if (i == 0) cycle
        count = count + 1
    end do
    print *, count
end program test
"#);
}

// ── Nested DO WHILE ───────────────────────────────────────────

#[test] fn nested_do_while() {
    compile_ok(r#"
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
"#);
}

#[test] fn nested_do_while_with_exit() {
    compile_ok(r#"
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
"#);
}

// ── DO WHILE with complex logical condition ───────────────────

#[test] fn do_while_and_condition() {
    compile_ok(r#"
program test
    integer :: x = 0, y = 10
    do while (x < 5 .and. y > 5)
        x = x + 1
        y = y - 1
    end do
    print *, x
    print *, y
end program test
"#);
}

#[test] fn do_while_or_condition() {
    compile_ok(r#"
program test
    integer :: a = 0, b = 0
    do while (a < 3 .or. b < 3)
        if (a < 3) a = a + 1
        if (b < 3) b = b + 1
    end do
    print *, a
    print *, b
end program test
"#);
}

#[test] fn do_while_not_condition() {
    compile_ok(r#"
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
"#);
}

// ── DO WHILE reading data ─────────────────────────────────────

#[test] fn do_while_read_until_eof() {
    compile_ok(r#"
program test
    integer :: n, ios, s
    s = 0
    do while (.true.)
        read(*, *, iostat=ios) n
        if (ios /= 0) exit
        s = s + n
    end do
    print *, s
end program test
"#);
}

// ── DO WHILE in subroutine ────────────────────────────────────

#[test] fn do_while_in_subroutine() {
    compile_ok(r#"
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
"#);
}

#[test] fn do_while_factorial() {
    compile_ok(r#"
program test
    integer :: n = 5, f = 1
    do while (n > 1)
        f = f * n
        n = n - 1
    end do
    print *, f
end program test
"#);
}

#[test] fn do_while_fibonacci() {
    compile_ok(r#"
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
"#);
}

#[test] fn do_while_newton_sqrt() {
    compile_ok(r#"
program test
    real :: x = 2.0, g = 1.0, prev
    do while (.true.)
        prev = g
        g = 0.5 * (g + x / g)
        if (abs(g - prev) < 1e-7) exit
    end do
    print *, abs(g * g - x) < 0.001
end program test
"#);
}

// ── Named DO WHILE (covered in test_named_loops but exercised here) ──

#[test] fn do_while_two_variables() {
    compile_ok(r#"
program test
    integer :: m = 1, n = 256
    integer :: steps = 0
    do while (m < n)
        m = m * 2
        steps = steps + 1
    end do
    print *, steps
end program test
"#);
}
