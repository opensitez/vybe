use super::helpers::run_prints;

#[test]
fn function_yield_returns_continuation() {
    let out = run_prints(
        r#"
program test
    print *, count()
contains
    function count() result(res)
        yield 1
    end function count
end program test
"#,
    );
    assert_eq!(out, vec!["[continuation]"]);
}
