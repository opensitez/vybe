use super::helpers::run_prints;

#[test]
fn select_case_complex_ranges_overlap_detection_with_preferred_order() {
    let out = run_prints(
        r#"
program select_case_complex_ranges_overlap_detection_with_preferred_order
    integer :: n
    n = 7
    select case (n)
    case (1:10)
        print *, 'first'
    case (5:8)
        print *, 'second'
    case default
        print *, 'third'
    end select
end program select_case_complex_ranges_overlap_detection_with_preferred_order
"#,
    );
    assert_eq!(out, vec!["first"]);
}

#[test]
fn select_case_complex_ranges_open_ended_lower() {
    let out = run_prints(
        r#"
program select_case_complex_ranges_open_ended_lower
    integer :: n
    n = -7
    select case (n)
    case (:-1)
        print *, 'neg'
    case (0:)
        print *, 'nonneg'
    end select
end program select_case_complex_ranges_open_ended_lower
"#,
    );
    assert_eq!(out, vec!["neg"]);
}

#[test]
fn select_case_complex_ranges_open_ended_upper() {
    let out = run_prints(
        r#"
program select_case_complex_ranges_open_ended_upper
    integer :: n
    n = 77
    select case (n)
    case (0:9)
        print *, 'single'
    case (10:99)
        print *, 'double'
    case (100:)
        print *, 'large'
    end select
end program select_case_complex_ranges_open_ended_upper
"#,
    );
    assert_eq!(out, vec!["double"]);
}

#[test]
fn select_case_complex_ranges_multiple_values_and_range_mix() {
    let out = run_prints(
        r#"
program select_case_complex_ranges_multiple_values_and_range_mix
    integer :: n
    n = 13
    select case (n)
    case (1, 3, 5, 9, 13)
        print *, 'odd-set'
    case (10:20)
        print *, 'middle'
    case default
        print *, 'out'
    end select
end program select_case_complex_ranges_multiple_values_and_range_mix
"#,
    );
    assert_eq!(out, vec!["odd-set"]);
}

#[test]
fn select_case_complex_ranges_guard_by_computed_bounds() {
    let out = run_prints(
        r#"
program select_case_complex_ranges_guard_by_computed_bounds
    integer :: n
    n = 42
    select case (n)
    case (1:4)
        print *, 'low'
    case (5:40)
        print *, 'mid'
    case (41:50)
        print *, 'high'
    case default
        print *, 'none'
    end select
end program select_case_complex_ranges_guard_by_computed_bounds
"#,
    );
    assert_eq!(out, vec!["high"]);
}

#[test]
fn select_case_complex_ranges_nested_case_in_loop() {
    let out = run_prints(
        r#"
program select_case_complex_ranges_nested_case_in_loop
    integer :: i
    do i = 1, 3
        select case (i)
        case (1)
            print *, 'one'
        case (2)
            print *, 'two'
        case default
            print *, 'more'
        end select
    end do
end program select_case_complex_ranges_nested_case_in_loop
"#,
    );
    assert_eq!(out, vec!["one", "two", "more"]);
}

#[test]
fn select_case_complex_ranges_character_like_with_ranges() {
    let out = run_prints(
        r#"
program select_case_complex_ranges_character_like_with_ranges
    character(len=1) :: c
    c = 'b'
    select case (c)
    case ('a':'d', 'f':'z')
        print *, 'group'
    case ('e')
        print *, 'alone'
    end select
end program select_case_complex_ranges_character_like_with_ranges
"#,
    );
    assert_eq!(out, vec!["group"]);
}

#[test]
fn select_case_complex_ranges_unmatched_value_default() {
    let out = run_prints(
        r#"
program select_case_complex_ranges_unmatched_value_default
    integer :: n
    n = 101
    select case (n)
    case (1:10)
        print *, 'low'
    case (11:20)
        print *, 'mid'
    case default
        print *, 'default'
    end select
end program select_case_complex_ranges_unmatched_value_default
"#,
    );
    assert_eq!(out, vec!["default"]);
}

#[test]
fn select_case_complex_ranges_real_like_integer_inputs() {
    let out = run_prints(
        r#"
program select_case_complex_ranges_real_like_integer_inputs
    integer :: n
    n = 2
    select case (n)
    case (:0)
        print *, 'nonpos'
    case (1, 3, 5)
        print *, 'odd-small'
    case (2, 4, 6)
        print *, 'even-small'
    case default
        print *, 'none'
    end select
end program select_case_complex_ranges_real_like_integer_inputs
"#,
    );
    assert_eq!(out, vec!["even-small"]);
}
