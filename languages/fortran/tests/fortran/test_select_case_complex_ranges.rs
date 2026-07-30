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
fn select_case_complex_ranges_no_default_unmatched_is_no_output() {
    let out = run_prints(
        r#"
program select_case_complex_ranges_no_default_unmatched_is_no_output
    integer :: n
    n = 0
    select case (n)
    case (1:10)
        print *, 'low'
    case (11:20)
        print *, 'mid'
    end select
end program select_case_complex_ranges_no_default_unmatched_is_no_output
"#,
    );
    assert_eq!(out, Vec::<&str>::new());
}

#[test]
fn select_case_complex_ranges_parameterized_range_bounds() {
    let out = run_prints(
        r#"
program select_case_complex_ranges_parameterized_range_bounds
    integer, parameter :: lo = 10
    integer, parameter :: hi = 20
    integer :: n
    n = 15
    select case (n)
    case (lo:hi)
        print *, 'parametric'
    case default
        print *, 'fallback'
    end select
end program select_case_complex_ranges_parameterized_range_bounds
"#,
    );
    assert_eq!(out, vec!["parametric"]);
}

#[test]
fn select_case_complex_ranges_overlap_between_list_and_range() {
    let out = run_prints(
        r#"
program select_case_complex_ranges_overlap_between_list_and_range
    integer :: n
    n = 4
    select case (n)
    case (1, 4, 7)
        print *, 'list'
    case (3:5)
        print *, 'range'
    case default
        print *, 'other'
    end select
end program select_case_complex_ranges_overlap_between_list_and_range
"#,
    );
    assert_eq!(out, vec!["list"]);
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

#[test]
fn select_case_complex_ranges_expression_selector() {
    let out = run_prints(
        r#"
program select_case_complex_ranges_expression_selector
    integer :: n
    n = 2
    select case (n + 1)
    case (1)
        print *, 'one'
    case (2:3)
        print *, 'plus'
    case (4)
        print *, 'four'
    end select
end program select_case_complex_ranges_expression_selector
"#,
    );
    assert_eq!(out, vec!["plus"]);
}

#[test]
fn select_case_complex_ranges_mixed_list_and_range_items() {
    let out = run_prints(
        r#"
program select_case_complex_ranges_mixed_list_and_range_items
    integer :: n
    n = 8
    select case (n)
    case (1, 4, 9, 12)
        print *, 'singles'
    case (6:8, 10:11, 14)
        print *, 'mixed'
    case default
        print *, 'other'
    end select
end program select_case_complex_ranges_mixed_list_and_range_items
"#,
    );
    assert_eq!(out, vec!["mixed"]);
}

#[test]
fn select_case_complex_ranges_character_overlap_prefers_first_case() {
    let out = run_prints(
        r#"
program select_case_complex_ranges_character_overlap_prefers_first_case
    character(len=1) :: c
    c = 'e'
    select case (c)
    case ('d':'h')
        print *, 'first'
    case ('a':'z')
        print *, 'second'
    case default
        print *, 'default'
    end select
end program select_case_complex_ranges_character_overlap_prefers_first_case
"#,
    );
    assert_eq!(out, vec!["first"]);
}

#[test]
fn select_case_complex_ranges_negative_open_ended_range_matches_only_negative() {
    let out = run_prints(
        r#"
program select_case_complex_ranges_negative_open_ended_range_matches_only_negative
    integer :: n
    n = -42
    select case (n)
    case (:-1)
        print *, 'negative'
    case (0:)
        print *, 'nonnegative'
    end select

    n = 42
    select case (n)
    case (:-1)
        print *, 'negative'
    case (0:)
        print *, 'nonnegative'
    end select
end program select_case_complex_ranges_negative_open_ended_range_matches_only_negative
"#,
    );
    assert_eq!(out, vec!["negative", "nonnegative"]);
}

#[test]
fn select_case_complex_ranges_zero_boundary_in_open_lower_range() {
    let out = run_prints(
        r#"
program select_case_complex_ranges_zero_boundary_in_open_lower_range
    integer :: n
    n = 0
    select case (n)
    case (:0)
        print *, 'lower-includes-zero'
    case (0:)
        print *, 'upper-includes-zero'
    case default
        print *, 'other'
    end select
end program select_case_complex_ranges_zero_boundary_in_open_lower_range
"#,
    );
    assert_eq!(out, vec!["lower-includes-zero"]);
}

#[test]
fn select_case_complex_ranges_singleton_precedes_range() {
    let out = run_prints(
        r#"
program select_case_complex_ranges_singleton_precedes_range
    integer :: n
    n = 6
    select case (n)
    case (6)
        print *, 'singleton'
    case (1:10)
        print *, 'range'
    case default
        print *, 'other'
    end select
end program select_case_complex_ranges_singleton_precedes_range
"#,
    );
    assert_eq!(out, vec!["singleton"]);
}

#[test]
fn select_case_complex_ranges_no_match_without_default_keeps_other_output_only() {
    let out = run_prints(
        r#"
program select_case_complex_ranges_no_match_without_default_keeps_other_output_only
    integer :: n
    n = 99
    select case (n)
    case (1:10)
        print *, 'low'
    case (20:30)
        print *, 'mid'
    end select
    print *, 'after'
end program select_case_complex_ranges_no_match_without_default_keeps_other_output_only
"#,
    );
    assert_eq!(out, vec!["after"]);
}
