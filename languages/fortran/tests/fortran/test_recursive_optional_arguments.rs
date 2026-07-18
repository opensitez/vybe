use super::helpers::run_prints;

#[test]
fn recursive_optional_arguments_default_chain() {
    let out = run_prints(
        r#"
program recursive_optional_arguments_default_chain
    print *, walk(3)
    print *, walk(3, 2)
contains
    recursive integer function walk(n, step) result(out)
        integer, intent(in) :: n
        integer, optional, intent(in) :: step
        integer :: stride
        if (present(step)) then
            stride = step
        else
            stride = 1
        end if
        if (n <= 0) then
            out = 0
        else
            out = n + walk(n - 1, stride)
        end if
    end function walk
end program recursive_optional_arguments_default_chain
"#,
    );
    assert_eq!(out, vec!["6", "9"]);
}

#[test]
fn recursive_optional_arguments_suffix_control() {
    let out = run_prints(
        r#"
program recursive_optional_arguments_suffix_control
    print *, fold(4)
    print *, fold(4, 2)
contains
    recursive integer function fold(n, step) result(out)
        integer, intent(in) :: n
        integer, optional, intent(in) :: step
        integer :: step_value
        step_value = 1
        if (present(step)) step_value = step
        if (n <= 0) then
            out = 0
        else
            out = n + fold(n - step_value, step_value)
        end if
    end function fold
end program recursive_optional_arguments_suffix_control
"#,
    );
    assert_eq!(out, vec!["10", "10"]);
}

#[test]
fn recursive_optional_arguments_character_default_behavior() {
    let out = run_prints(
        r#"
program recursive_optional_arguments_character_default_behavior
    print *, letters('b')
    print *, letters('b', 2)
contains
    recursive character(len=16) function letters(ch, repeat_count) result(out)
        character(len=*), intent(in) :: ch
        integer, optional, intent(in) :: repeat_count
        if (present(repeat_count) .and. repeat_count > 0) then
            if (len_trim(ch) >= 2) then
                out = ch // '_'
            else
                out = trim(letters(ch, repeat_count - 1))
            end if
        else
            out = ch
        end if
    end function letters
end program recursive_optional_arguments_character_default_behavior
"#,
    );
    assert_eq!(out, vec!["b", "b"]);
}

#[test]
fn recursive_optional_arguments_optional_limit_guard() {
    let out = run_prints(
        r#"
program recursive_optional_arguments_optional_limit_guard
    print *, limited(5)
    print *, limited(5, 2)
contains
    recursive integer function limited(n, limit) result(out)
        integer, intent(in) :: n
        integer, optional, intent(in) :: limit
        integer :: max_depth
        max_depth = 1
        if (present(limit)) max_depth = limit
        if (n <= 0 .or. n > 10*max_depth) then
            out = 0
        else
            out = 1 + limited(n - 1, max_depth)
        end if
    end function limited
end program recursive_optional_arguments_optional_limit_guard
"#,
    );
    assert_eq!(out, vec!["6", "6"]);
}

#[test]
fn recursive_optional_arguments_optional_direction_signals() {
    let out = run_prints(
        r#"
program recursive_optional_arguments_optional_direction_signals
    print *, sign_walk(4)
    print *, sign_walk(4, -1)
contains
    recursive integer function sign_walk(n, dir) result(out)
        integer, intent(in) :: n
        integer, optional, intent(in) :: dir
        integer :: d
        d = 1
        if (present(dir)) d = dir
        if (abs(d) /= 1) d = 1
        if (n <= 0) then
            out = 0
        else if (d == -1) then
            out = 1 + sign_walk(n - 1)
        else
            out = n
        end if
    end function sign_walk
end program recursive_optional_arguments_optional_direction_signals
"#,
    );
    assert_eq!(out, vec!["4", "4"]);
}

#[test]
fn recursive_optional_arguments_optional_weighted_accumulate() {
    let out = run_prints(
        r#"
program recursive_optional_arguments_optional_weighted_accumulate
    print *, weighted(3)
    print *, weighted(3, 10)
contains
    recursive integer function weighted(n, weight) result(out)
        integer, intent(in) :: n
        integer, optional, intent(in) :: weight
        integer :: w
        w = 1
        if (present(weight)) w = weight
        if (n <= 0) then
            out = 0
        else
            out = n * w + weighted(n - 1)
        end if
    end function weighted
end program recursive_optional_arguments_optional_weighted_accumulate
"#,
    );
    assert_eq!(out, vec!["6", "33"]);
}

#[test]
fn recursive_optional_arguments_logical_optional_termination() {
    let out = run_prints(
        r#"
program recursive_optional_arguments_logical_optional_termination
    print *, parity(6)
    print *, parity(6, .false.)
contains
    recursive integer function parity(n, even_only) result(out)
        integer, intent(in) :: n
        logical, optional, intent(in) :: even_only
        logical :: only_even
        only_even = .true.
        if (present(even_only)) only_even = even_only
        if (n <= 0) then
            out = 0
        else if (mod(n,2) == 0 .or. .not. only_even) then
            out = n + parity(n - 1, only_even)
        else
            out = parity(n - 1, only_even)
        end if
    end function parity
end program recursive_optional_arguments_logical_optional_termination
"#,
    );
    assert_eq!(out, vec!["21", "9"]);
}

#[test]
fn recursive_optional_arguments_optional_return_zero_path() {
    let out = run_prints(
        r#"
program recursive_optional_arguments_optional_return_zero_path
    print *, capped(9)
    print *, capped(9, 4)
contains
    recursive integer function capped(n, cap) result(out)
        integer, intent(in) :: n
        integer, optional, intent(in) :: cap
        integer :: m
        if (present(cap)) m = cap
        if (n <= 0) then
            out = 0
        else if (present(cap) .and. n > m) then
            out = m
        else
            out = n + capped(n - 1, cap)
        end if
    end function capped
end program recursive_optional_arguments_optional_return_zero_path
"#,
    );
    assert_eq!(out, vec!["45", "10"]);
}
