use super::helpers::run_prints;

#[test]
fn test_do_construct_reentrancy_nested_accumulators() {
    let out = run_prints(
        r#"
program test_do_construct_reentrancy
    integer :: outer, inner, total
    total = 0
    do outer = 1, 3
        do inner = 1, 2
            total = total + outer * inner
        end do
    end do
    print *, total
end program test_do_construct_reentrancy
"#,
    );

    assert_eq!(out, vec!["18"]);
}
