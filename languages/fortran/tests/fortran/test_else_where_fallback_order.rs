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
