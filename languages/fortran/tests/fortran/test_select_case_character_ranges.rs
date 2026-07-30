use super::helpers::run_prints;

#[test]
fn select_case_character_ranges_single_range() {
    let out = run_prints(
        r#"
program select_case_character_ranges_single_range
    character(len=1) :: c
    c = 'f'
    select case (c)
    case ('a':'m')
        print *, 'low'
    case ('n':'z')
        print *, 'high'
    end select
end program select_case_character_ranges_single_range
"#,
    );
    assert_eq!(out, vec!["low"]);
}

#[test]
fn select_case_character_ranges_ascii_boundary() {
    let out = run_prints(
        r#"
program select_case_character_ranges_ascii_boundary
    character(len=1) :: c
    c = 'z'
    select case (c)
    case ('a':'m')
        print *, 'A'
    case ('n':'z')
        print *, 'Z'
    case default
        print *, 'D'
    end select
end program select_case_character_ranges_ascii_boundary
"#,
    );
    assert_eq!(out, vec!["Z"]);
}

#[test]
fn select_case_character_ranges_multiple_ranges_on_one_case() {
    let out = run_prints(
        r#"
program select_case_character_ranges_multiple_ranges_on_one_case
    character(len=1) :: c
    c = 'k'
    select case (c)
    case ('a':'c', 'k':'m')
        print *, 'ok'
    case default
        print *, 'no'
    end select
end program select_case_character_ranges_multiple_ranges_on_one_case
"#,
    );
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn select_case_character_ranges_empty_string_case() {
    let out = run_prints(
        r#"
program select_case_character_ranges_empty_string_case
    character(len=3) :: c
    c = ''
    select case (c)
    case ('')
        print *, 'empty'
    case default
        print *, 'non-empty'
    end select
end program select_case_character_ranges_empty_string_case
"#,
    );
    assert_eq!(out, vec!["empty"]);
}

#[test]
fn select_case_character_ranges_trimmed_length_case() {
    let out = run_prints(
        r#"
program select_case_character_ranges_trimmed_length_case
    character(len=6) :: c
    c = 'abc   '
    select case (trim(c))
    case ('abc')
        print *, 'trimmed'
    case default
        print *, 'raw'
    end select
end program select_case_character_ranges_trimmed_length_case
"#,
    );
    assert_eq!(out, vec!["trimmed"]);
}

#[test]
fn select_case_character_ranges_numeric_text_tokens() {
    let out = run_prints(
        r#"
program select_case_character_ranges_numeric_text_tokens
    character(len=4) :: c
    c = 'v2x'
    select case (c)
    case ('a':'m')
        print *, 'alpha'
    case ('v1':'v9')
        print *, 'version'
    case default
        print *, 'other'
    end select
end program select_case_character_ranges_numeric_text_tokens
"#,
    );
    assert_eq!(out, vec!["version"]);
}

#[test]
fn select_case_character_ranges_word_ranges() {
    let out = run_prints(
        r#"
program select_case_character_ranges_word_ranges
    character(len=5) :: c
    c = 'beta '
    select case (trim(c))
    case ('alpha')
        print *, 'first'
    case ('beta', 'gamma')
        print *, 'second'
    case ('delta')
        print *, 'third'
    case default
        print *, 'else'
    end select
end program select_case_character_ranges_word_ranges
"#,
    );
    assert_eq!(out, vec!["second"]);
}

#[test]
fn select_case_character_ranges_pure_default_branch() {
    let out = run_prints(
        r#"
program select_case_character_ranges_pure_default_branch
    character(len=3) :: c
    c = 'xyz'
    select case (c)
    case ('a':'f')
        print *, 'small'
    case default
        print *, 'default'
    end select
end program select_case_character_ranges_pure_default_branch
"#,
    );
    assert_eq!(out, vec!["default"]);
}

#[test]
fn select_case_character_ranges_overlap_respects_first_case() {
    let out = run_prints(
        r#"
program select_case_character_ranges_overlap_respects_first_case
    character(len=1) :: c
    c = 'f'
    select case (c)
    case ('a':'z')
        print *, 'all'
    case ('d':'h')
        print *, 'subset'
    case default
        print *, 'default'
    end select
end program select_case_character_ranges_overlap_respects_first_case
"#,
    );
    assert_eq!(out, vec!["all"]);
}

#[test]
fn select_case_character_ranges_boundary_edge() {
    let out = run_prints(
        r#"
program select_case_character_ranges_boundary_edge
    character(len=1) :: c
    c = 'n'
    select case (c)
    case ('a':'m')
        print *, 'low'
    case ('n':'z')
        print *, 'high'
    case default
        print *, 'fallback'
    end select
end program select_case_character_ranges_boundary_edge
"#,
    );
    assert_eq!(out, vec!["high"]);
}

#[test]
fn select_case_character_ranges_mixed_values_and_range() {
    let out = run_prints(
        r#"
program select_case_character_ranges_mixed_values_and_range
    character(len=1) :: c
    c = 'k'
    select case (c)
    case ('a', 'c', 'k', 'z')
        print *, 'list-hit'
    case ('f':'t')
        print *, 'range-hit'
    case default
        print *, 'none'
    end select
end program select_case_character_ranges_mixed_values_and_range
"#,
    );
    assert_eq!(out, vec!["list-hit"]);
}

#[test]
fn select_case_character_ranges_digit_range() {
    let out = run_prints(
        r#"
program select_case_character_ranges_digit_range
    character(len=1) :: c
    c = '9'
    select case (c)
    case ('0':'3')
        print *, 'small'
    case ('4':'9')
        print *, 'large'
    case ('a':'z')
        print *, 'letter'
    case default
        print *, 'other'
    end select
end program select_case_character_ranges_digit_range
"#,
    );
    assert_eq!(out, vec!["large"]);
}

#[test]
fn select_case_character_ranges_open_ended_lower_bound() {
    let out = run_prints(
        r#"
program select_case_character_ranges_open_ended_lower_bound
    character(len=1) :: c
    c = 'e'
    select case (c)
    case (:'f')
        print *, 'low-half'
    case ('g':)
        print *, 'high-half'
    end select
end program select_case_character_ranges_open_ended_lower_bound
"#,
    );
    assert_eq!(out, vec!["low-half"]);
}

#[test]
fn select_case_character_ranges_open_ended_upper_bound() {
    let out = run_prints(
        r#"
program select_case_character_ranges_open_ended_upper_bound
    character(len=1) :: c
    c = 'p'
    select case (c)
    case (:'k')
        print *, 'low'
    case ('l':)
        print *, 'high'
    end select
end program select_case_character_ranges_open_ended_upper_bound
"#,
    );
    assert_eq!(out, vec!["high"]);
}

#[test]
fn select_case_character_ranges_implicit_overlap_prefers_first() {
    let out = run_prints(
        r#"
program select_case_character_ranges_implicit_overlap_prefers_first
    character(len=1) :: c
    c = 'a'
    select case (c)
    case ('a':'m')
        print *, 'prefix'
    case ('b':)
        print *, 'fallback'
    case default
        print *, 'none'
    end select
end program select_case_character_ranges_implicit_overlap_prefers_first
    "#,
    );
    assert_eq!(out, vec!["prefix"]);
}

#[test]
fn select_case_character_ranges_no_default_no_match() {
    let out = run_prints(
        r#"
program select_case_character_ranges_no_default_no_match
    character(len=2) :: c
    c = 'zZ'
    select case (c)
    case ('a')
        print *, 'alpha'
    case ('b')
        print *, 'bravo'
    end select
    print *, 'done'
end program select_case_character_ranges_no_default_no_match
"#,
    );
    assert_eq!(out, vec!["done"]);
}

#[test]
fn select_case_character_ranges_open_range_lower_bound_edge() {
    let out = run_prints(
        r#"
program select_case_character_ranges_open_range_lower_bound_edge
    character(len=1) :: c
    c = 'a'
    select case (c)
    case (:'m')
        print *, 'first'
    case ('n':)
        print *, 'second'
    case default
        print *, 'other'
    end select
end program select_case_character_ranges_open_range_lower_bound_edge
"#,
    );
    assert_eq!(out, vec!["first"]);
}

#[test]
fn select_case_character_ranges_multi_value_exact_precedence() {
    let out = run_prints(
        r#"
program select_case_character_ranges_multi_value_exact_precedence
    character(len=1) :: c
    c = 'k'
    select case (c)
    case ('a', 'k', 'z')
        print *, 'list-hit'
    case ('d':'m')
        print *, 'range-hit'
    case default
        print *, 'none'
    end select
end program select_case_character_ranges_multi_value_exact_precedence
"#,
    );
    assert_eq!(out, vec!["list-hit"]);
}
