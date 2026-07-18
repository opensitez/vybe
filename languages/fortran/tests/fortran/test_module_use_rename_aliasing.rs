use super::helpers::run_prints;

#[test]
fn test_module_use_rename_aliasing_rebinds_symbol() {
    let out = run_prints(
        r#"
module alias_mod
    integer, parameter :: original = 21
end module

program test_module_use_rename_aliasing
    use alias_mod, only: short => original
    print *, short
end program test_module_use_rename_aliasing
"#,
    );

    assert_eq!(out, vec!["21"]);
}
