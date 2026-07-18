use super::helpers::run_prints;

#[test]
fn test_io_list_directed_record_gaps_parse_with_commas() {
    let out = run_prints(
        r#"
program test_io_list_directed_record_gaps
    character(len=80) :: text
    integer :: a
    integer :: b
    text = '1, 2, 3'
    read(text, *) a
    read(text(4:80), *) b
    print *, a + b
end program test_io_list_directed_record_gaps
"#,
    );

    assert_eq!(out, vec!["3"]);
}
