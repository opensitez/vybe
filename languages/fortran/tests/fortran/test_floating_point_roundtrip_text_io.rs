use super::helpers::run_prints;

#[test]
fn test_floating_point_roundtrip_text_io_parse_and_format() {
    let out = run_prints(
        r#"
program test_floating_point_roundtrip_text_io
    real :: value
    character(len=20) :: text
    write(text, '(F8.3)') 1.5
    read(text, '(F8.3)') value
    print *, value
end program test_floating_point_roundtrip_text_io
"#,
    );

    assert_eq!(out, vec!["1.5"]);
}
