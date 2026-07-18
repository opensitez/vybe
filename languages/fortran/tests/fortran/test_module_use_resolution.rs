use super::helpers::run_prints;

#[test]
fn test_module_use_resolution_pulls_public_binding() {
    let out = run_prints(
        r#"
module math_constants
    integer, parameter :: magic = 11
end module

program test_module_use_resolution
    use math_constants
    print *, magic
end program test_module_use_resolution
"#,
    );

    assert_eq!(out, vec!["11"]);
}
