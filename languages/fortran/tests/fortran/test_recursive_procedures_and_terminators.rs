use super::helpers::run_prints;

#[test]
fn recursive_procedures_and_terminators_factorial_guard() {
    let out = run_prints(
        r#"
program recursive_procedures_and_terminators_factorial_guard
    print *, fact(0)
    print *, fact(5)
contains
    recursive integer function fact(n) result(out)
        integer, intent(in) :: n
        if (n <= 1) then
            out = 1
        else
            out = n * fact(n - 1)
        end if
    end function fact
end program recursive_procedures_and_terminators_factorial_guard
"#,
    );
    assert_eq!(out, vec!["1", "120"]);
}

#[test]
fn recursive_procedures_and_terminators_mutual_recursion_odd_even() {
    let out = run_prints(
        r#"
program recursive_procedures_and_terminators_mutual_recursion_odd_even
    print *, is_even(4)
    print *, is_odd(4)
contains
    recursive logical function is_even(n)
        integer, intent(in) :: n
        if (n == 0) then
            is_even = .true.
        else
            is_even = is_odd(n - 1)
        end if
    end function is_even

    recursive logical function is_odd(n)
        integer, intent(in) :: n
        if (n == 0) then
            is_odd = .false.
        else
            is_odd = is_even(n - 1)
        end if
    end function is_odd
end program recursive_procedures_and_terminators_mutual_recursion_odd_even
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn recursive_procedures_and_terminators_terminator_after_if() {
    let out = run_prints(
        r#"
program recursive_procedures_and_terminators_terminator_after_if
    integer :: sum
    sum = accumulator(1)
    print *, sum
contains
    recursive integer function accumulator(n) result(out)
        integer, intent(in) :: n
        if (n >= 6) then
            out = 0
        else
            out = n + accumulator(n + 1)
        end if
    end function accumulator
end program recursive_procedures_and_terminators_terminator_after_if
"#,
    );
    assert_eq!(out, vec!["21"]);
}

#[test]
fn recursive_procedures_and_terminators_early_return_pattern() {
    let out = run_prints(
        r#"
program recursive_procedures_and_terminators_early_return_pattern
    print *, first_nonzero((/0, 0, 5, 7/), 1)
contains
    recursive integer function first_nonzero(values, idx) result(out)
        integer, intent(in) :: values(:)
        integer, intent(in) :: idx
        if (idx > size(values)) then
            out = -1
        else if (values(idx) /= 0) then
            out = values(idx)
        else
            out = first_nonzero(values, idx + 1)
        end if
    end function first_nonzero
end program recursive_procedures_and_terminators_early_return_pattern
"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn recursive_procedures_and_terminators_terminator_via_zero_stride_guard() {
    let out = run_prints(
        r#"
program recursive_procedures_and_terminators_terminator_via_zero_stride_guard
    print *, countdown(3)
contains
    recursive integer function countdown(n) result(out)
        integer, intent(in) :: n
        if (n <= 0) then
            out = 0
        else
            out = 1 + countdown(n - 1)
        end if
    end function countdown
end program recursive_procedures_and_terminators_terminator_via_zero_stride_guard
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn recursive_procedures_and_terminators_tail_guarding_depth() {
    let out = run_prints(
        r#"
program recursive_procedures_and_terminators_tail_guarding_depth
    print *, depth_walk(4)
contains
    recursive integer function depth_walk(n) result(out)
        integer, intent(in) :: n
        if (n <= 0) then
            out = 0
        else
            out = depth_walk(n - 2) + 1
        end if
    end function depth_walk
end program recursive_procedures_and_terminators_tail_guarding_depth
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn recursive_procedures_and_terminators_return_accumulator() {
    let out = run_prints(
        r#"
program recursive_procedures_and_terminators_return_accumulator
    integer :: total
    total = series(1, 4)
    print *, total
contains
    recursive integer function series(a, b) result(out)
        integer, intent(in) :: a, b
        if (a > b) then
            out = 0
        else
            out = a + series(a + 1, b)
        end if
    end function series
end program recursive_procedures_and_terminators_return_accumulator
"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn recursive_procedures_and_terminators_branching_termination() {
    let out = run_prints(
        r#"
program recursive_procedures_and_terminators_branching_termination
    print *, sum_until(1, 10)
    print *, sum_until(1, 2)
contains
    recursive integer function sum_until(start, limit) result(out)
        integer, intent(in) :: start
        integer, intent(in) :: limit
        if (start > limit) then
            out = 0
        else if (start == limit) then
            out = start
        else
            out = start + sum_until(start + 1, limit)
        end if
    end function sum_until
end program recursive_procedures_and_terminators_branching_termination
"#,
    );
    assert_eq!(out, vec!["55", "3"]);
}
