use super::helpers::run_prints;

#[test]
fn test_io_pure_output_buffering_prints_as_expected() {
    let out = run_prints(
        r#"
program test_io_pure_output_buffering
    print *, 1
    print *, 2
    print *, 3
end program test_io_pure_output_buffering
"#,
    );

    assert_eq!(out, vec!["1", "2", "3"]);
}
