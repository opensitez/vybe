use super::helpers::run_prints;

// ── Basic named DO loop ───────────────────────────────────────

#[test]
fn named_do_basic() {
    let out = run_prints(
        r#"
program test
    integer :: i, s
    s = 0
    outer: do i = 1, 5
        s = s + i
    end do outer
    print *, s
end program test
"#,
    );

    assert_eq!(out, vec!["15"]);
}

#[test]
fn named_do_while() {
    let out = run_prints(
        r#"
program test
    integer :: n = 0
    counting: do while (n < 5)
        n = n + 1
    end do counting
    print *, n
end program test
"#,
    );

    assert_eq!(out, vec!["5"]);
}

#[test]
fn named_do_nested_both() {
    let out = run_prints(
        r#"
program test
    integer :: i, j, s
    s = 0
    outer: do i = 1, 3
        inner: do j = 1, 3
            s = s + 1
        end do inner
    end do outer
    print *, s
end program test
"#,
    );

    assert_eq!(out, vec!["9"]);
}

// ── EXIT with loop name ───────────────────────────────────────

#[test]
fn exit_named_outer() {
    let out = run_prints(
        r#"
program test
    integer :: i, j
    outer: do i = 1, 5
        inner: do j = 1, 5
            if (i == 3 .and. j == 3) exit outer
        end do inner
    end do outer
    print *, i
    print *, j
end program test
"#,
    );

    assert_eq!(out, vec!["3", "3"]);
}

#[test]
fn exit_named_inner() {
    let out = run_prints(
        r#"
program test
    integer :: i, j, count
    count = 0
    outer: do i = 1, 3
        inner: do j = 1, 10
            if (j > 3) exit inner
            count = count + 1
        end do inner
    end do outer
    print *, count
end program test
"#,
    );

    assert_eq!(out, vec!["9"]);
}

#[test]
fn exit_outer_from_deep_nest() {
    let out = run_prints(
        r#"
program test
    integer :: i, j, k
    outer: do i = 1, 10
        middle: do j = 1, 10
            inner: do k = 1, 10
                if (i + j + k == 10) exit outer
            end do inner
        end do middle
    end do outer
    print *, i, j, k
end program test
"#,
    );

    assert_eq!(out, vec!["2", "1", "7"]);
}

#[test]
fn exit_named_vs_unnamed() {
    let out = run_prints(
        r#"
program test
    integer :: i, j
    named: do i = 1, 5
        do j = 1, 5
            if (j == 3) exit named
        end do
        print *, i
    end do named
    print *, 'done'
end program test
"#,
    );
    assert_eq!(out, vec!["1", "done"]);
}

// ── CYCLE with loop name ──────────────────────────────────────

#[test]
fn cycle_named_outer() {
    let out = run_prints(
        r#"
program test
    integer :: i, j, count
    count = 0
    outer: do i = 1, 5
        inner: do j = 1, 5
            if (j == 3) cycle outer
            count = count + 1
        end do inner
    end do outer
    print *, count
end program test
"#,
    );

    assert_eq!(out, vec!["10"]);
}

#[test]
fn cycle_named_inner() {
    let out = run_prints(
        r#"
program test
    integer :: i, j, count
    count = 0
    outer: do i = 1, 3
        inner: do j = 1, 5
            if (j == 3) cycle inner
            count = count + 1
        end do inner
    end do outer
    print *, count
end program test
"#,
    );

    assert_eq!(out, vec!["12"]);
}

#[test]
fn cycle_outer_skip_rest_of_inner() {
    let out = run_prints(
        r#"
program test
    integer :: i, j
    outer: do i = 1, 3
        inner: do j = 1, 4
            if (j == 2) cycle outer
            print *, i, j
        end do inner
    end do outer
end program test
"#,
    );
    assert_eq!(out, vec!["1 1", "2 1", "3 1"]);
}

#[test]
fn cycle_preserves_state() {
    let out = run_prints(
        r#"
program test
    integer :: i, j, sum_j
    sum_j = 0
    outer: do i = 1, 4
        inner: do j = 1, 4
            if (mod(j, 2) == 0) cycle outer
            sum_j = sum_j + j
        end do inner
    end do outer
    print *, sum_j
end program test
"#,
    );
    assert_eq!(out, vec!["16"]);
}

// ── Three-level named loops ───────────────────────────────────

#[test]
fn three_level_exit_outer() {
    let out = run_prints(
        r#"
program test
    integer :: i, j, k
    found: do i = 1, 10
        mid: do j = 1, 10
            deep: do k = 1, 10
                if (i * j * k == 24) exit found
            end do deep
        end do mid
    end do found
    print *, i * j * k
end program test
"#,
    );
    assert_eq!(out, vec!["24"]);
}

#[test]
fn three_level_cycle_middle() {
    let out = run_prints(
        r#"
program test
    integer :: i, j, k, count
    count = 0
    outer: do i = 1, 3
        mid: do j = 1, 3
            inner: do k = 1, 3
                if (k == 2) cycle mid
                count = count + 1
            end do inner
        end do mid
    end do outer
    print *, count
end program test
"#,
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn three_level_exit_middle() {
    let out = run_prints(
        r#"
program test
    integer :: i, j, k, count
    count = 0
    outer: do i = 1, 4
        mid: do j = 1, 4
            inner: do k = 1, 4
                if (j == 2 .and. k == 2) exit mid
                count = count + 1
            end do inner
        end do mid
    end do outer
    print *, count
end program test
"#,
    );
    assert_eq!(out, vec!["16"]);
}

// ── Named DO CONCURRENT ───────────────────────────────────────

#[test]
fn named_do_concurrent() {
    let out = run_prints(
        r#"
program test
    integer :: a(10)
    fill: do concurrent (i = 1:10)
        a(i) = i * i
    end do fill
    print *, a(5)
end program test
"#,
    );
    assert_eq!(out, vec!["25"]);
}

// ── Named loops with subroutine calls inside ─────────────────

#[test]
fn named_loop_with_call() {
    let out = run_prints(
        r#"
program test
    integer :: i, total
    total = 0
    accumulate: do i = 1, 10
        if (mod(i, 3) == 0) cycle accumulate
        call add(total, i)
    end do accumulate
    print *, total
contains
    subroutine add(acc, n)
        integer, intent(inout) :: acc
        integer, intent(in)    :: n
        acc = acc + n
end subroutine add
end program test
"#,
    );
    assert_eq!(out, vec!["37"]);
}

// ── Edge cases ────────────────────────────────────────────────

#[test]
fn named_loop_zero_iterations() {
    let out = run_prints(
        r#"
program test
    integer :: i
    nothing: do i = 5, 1
        print *, i
    end do nothing
    print *, 'done'
end program test
"#,
    );

    assert_eq!(out, vec!["done"]);
}

#[test]
fn exit_at_start_of_loop() {
    let out = run_prints(
        r#"
program test
    integer :: i
    quick: do i = 1, 100
        exit quick
    end do quick
    print *, i
end program test
"#,
    );

    assert_eq!(out, vec!["1"]);
}

#[test]
fn nested_exit_and_cycle_mix() {
    let out = run_prints(
        r#"
program test
    integer :: i, j, s
    s = 0
    outer: do i = 1, 5
        inner: do j = 1, 5
            if (j == 1) cycle inner
            if (j == 4) cycle outer
            if (i == 4) exit outer
            s = s + 1
        end do inner
    end do outer
    print *, s
end program test
"#,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn named_do_named_loop_value_accumulates() {
    let out = run_prints(
        r#"
program test
    integer :: i, s
    s = 0
    outer: do i = 1, 4
        s = s + i
    end do outer
    print *, s
end program test
"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn named_do_named_while_runs_and_leaves_value() {
    let out = run_prints(
        r#"
program test
    integer :: n
    integer :: count
    n = 0
    count = 0
    counting: do while (n < 3)
        n = n + 1
        count = count + n
    end do counting
    print *, count
end program test
"#,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn named_do_outer_cycle_skips_to_next_outer_iteration() {
    let out = run_prints(
        r#"
program test
    integer :: i, j, s
    s = 0
    outer: do i = 1, 3
        inner: do j = 1, 4
            if (j == 2) cycle outer
            s = s + 1
        end do inner
    end do outer
    print *, s
end program test
"#,
    );
    assert_eq!(out, vec!["3"]);
}
