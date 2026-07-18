use super::helpers::run_prints;

#[test]
fn test_else_if_cascade_priority_resolves_first_match() {
    let out = run_prints(
        r#"
program test_else_if_cascade_priority
    integer :: x
    x = 3
    if (x > 4) then
        print *, 1
    else if (x == 3) then
        print *, 2
    else
        print *, 3
    end if
end program test_else_if_cascade_priority
"#,
    );

    assert_eq!(out, vec!["2"]);
}
