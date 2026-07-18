use super::helpers::run_prints;

#[test]
fn test_name_resolution_renamed_imports_prefers_renamed_symbol() {
    let out = run_prints(
        r#"
module source_mod
    integer, parameter :: base = 14
end module

program test_name_resolution_renamed_imports
    use source_mod, only: exported => base
    print *, exported
end program test_name_resolution_renamed_imports
"#,
    );

    assert_eq!(out, vec!["14"]);
}
