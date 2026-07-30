use super::helpers::run_prints;

#[test]
fn test_else_where_fallback_order_applies_primary_mask() {
    let out = run_prints(
        r#"
program test_else_where_fallback_order
    integer :: a(4)
    integer :: b(4)
    a = (/1, 2, 3, 4/)
    b = 0
    where (a > 2)
        b = 1
    elsewhere where (a == 2)
        b = 2
    elsewhere
        b = 3
    end where
    print *, b(1)
    print *, b(2)
    print *, b(3)
    print *, b(4)
end program test_else_where_fallback_order
"#,
    );

    assert_eq!(out, vec!["3", "2", "1", "1"]);
}

#[test]
fn test_else_where_default_only_runs_when_previous_masks_fail() {
    let out = run_prints(
        r#"
program test_else_where_fallback_default_only
    integer :: a(4)
    integer :: b(4)
    a = (/0, 1, 2, 3/)
    b = 0
    where (a == 10)
        b = 1
    elsewhere where (a == 20)
        b = 2
    elsewhere
        b = 9
    end where
    print *, b(1)
    print *, b(2)
    print *, b(3)
    print *, b(4)
end program test_else_where_fallback_default_only
"#,
    );

    assert_eq!(out, vec!["9", "9", "9", "9"]);
}

#[test]
fn test_else_where_with_multiple_elsewhere_sections() {
    let out = run_prints(
        r#"
program test_else_where_fallback_multiple_sections
    integer :: a(5)
    integer :: b(5)
    a = (/1, 2, 3, 4, 5/)
    b = 0
    where (a <= 2)
        b = 5
    elsewhere where (a > 3)
        b = 10
    elsewhere where (a == 3)
        b = 7
    elsewhere
        b = 99
    end where
    print *, b(1)
    print *, b(2)
    print *, b(3)
    print *, b(4)
    print *, b(5)
end program test_else_where_fallback_multiple_sections
"#,
    );

    assert_eq!(out, vec!["5", "5", "7", "10", "10"]);
}

#[test]
fn test_else_where_order_prefers_first_true_mask() {
    let out = run_prints(
        r#"
program test_else_where_order_prefers_first
    integer :: a(4)
    integer :: b(4)
    a = (/1, 2, 3, 4/)
    b = 0
    where (a > 0)
        b = 1
    elsewhere where (a < 4)
        b = 2
    elsewhere
        b = 3
    end where
    print *, b(1)
    print *, b(2)
    print *, b(3)
    print *, b(4)
end program test_else_where_order_prefers_first
"#,
    );

    assert_eq!(out, vec!["1", "1", "1", "1"]);
}

#[test]
fn test_else_where_masked_even_odd_fallthrough() {
    let out = run_prints(
        r#"
program test_else_where_masked_even_odd_fallthrough
    integer :: a(5)
    integer :: b(5)
    a = (/1, 2, 3, 4, 5/)
    b = 0
    where (a > 3)
        b = 10
    elsewhere where (mod(a, 2) == 1)
        b = 20
    elsewhere
        b = 30
    end where
    print *, b(1)
    print *, b(2)
    print *, b(3)
    print *, b(4)
    print *, b(5)
end program test_else_where_masked_even_odd_fallthrough
"#,
    );

    assert_eq!(out, vec!["20", "30", "20", "10", "10"]);
}

#[test]
fn test_else_where_chain_skips_unreached_sections() {
    let out = run_prints(
        r#"
program test_else_where_chain_skips_unreached
    integer :: a(5)
    integer :: b(5)
    a = (/1, 2, 3, 4, 5/)
    b = 0
    where (a > 10)
        b = 1
    elsewhere (a > 3)
        b = 2
    elsewhere (a == 4)
        b = 9
    elsewhere
        b = 3
    end where
    print *, b(1)
    print *, b(4)
    print *, b(5)
end program test_else_where_chain_skips_unreached
"#,
    );

    assert_eq!(out, vec!["3", "2", "3"]);
}

#[test]
fn test_else_where_all_true_first_clause() {
    let out = run_prints(
        r#"
program test_else_where_all_true_first_clause
    integer :: a(4)
    integer :: b(4)
    a = (/10, 20, 30, 40/)
    where (a > 0)
        b = 5
    elsewhere
        b = 9
    end where
    print *, b(1)
    print *, b(4)
end program test_else_where_all_true_first_clause
"#,
    );

    assert_eq!(out, vec!["5", "5"]);
}

#[test]
fn test_else_where_scalar_match() {
    let out = run_prints(
        r#"
program test_else_where_scalar_match
    integer :: a(1)
    integer :: b(1)
    a = (/7/)
    where (a == 7)
        b = 77
    elsewhere
        b = 13
    end where
    print *, b(1)
end program test_else_where_scalar_match
"#,
    );

    assert_eq!(out, vec!["77"]);
}

#[test]
fn test_else_where_no_default_no_match_preserves_existing_values() {
    let out = run_prints(
        r#"
program test_else_where_no_default_no_match_preserves
    integer :: a(4)
    integer :: b(4)
    a = (/1, 2, 3, 4/)
    b = (/10, 20, 30, 40/)
    where (a > 4)
        b = 10 * a
    elsewhere (a < 0)
        b = -10
    end where
    print *, b(1)
    print *, b(2)
    print *, b(3)
    print *, b(4)
end program test_else_where_no_default_no_match_preserves
"#,
    );

    assert_eq!(out, vec!["10", "20", "30", "40"]);
}

#[test]
fn test_else_where_single_match_applies_before_no_default_mask() {
    let out = run_prints(
        r#"
program test_else_where_single_match_applies
    integer :: a(3)
    integer :: b(3)
    a = (/1, 5, 9/)
    b = (/7, 8, 9/)
    where (a == 5)
        b = 50
    elsewhere (a == 6)
        b = 60
    end where
    print *, b(1)
    print *, b(2)
    print *, b(3)
end program test_else_where_single_match_applies
"#,
    );

    assert_eq!(out, vec!["7", "50", "9"]);
}
