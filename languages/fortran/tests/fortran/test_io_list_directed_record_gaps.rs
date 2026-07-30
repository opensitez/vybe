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

#[test]
fn test_io_list_directed_record_gaps_handles_multiple_whitespace() {
    let out = run_prints(
        r#"
program test_io_list_directed_record_gaps
    character(len=80) :: text
    integer :: a, b, c
    text = '  10    20   30  '
    read(text, *) a, b, c
    print *, a + b + c
end program test_io_list_directed_record_gaps
"#,
    );

    assert_eq!(out, vec!["60"]);
}

#[test]
fn test_io_list_directed_record_gaps_parse_mixed_types() {
    let out = run_prints(
        r#"
program test_io_list_directed_record_gaps
    character(len=80) :: text
    integer :: a
    real :: x
    logical :: f
    text = '42 3.5 .true.'
    read(text, *) a, x, f
    print *, a
    print *, int(x)
    print *, f
end program test_io_list_directed_record_gaps
"#,
    );

    assert_eq!(out, vec!["42", "3", "true"]);
}

#[test]
fn test_io_list_directed_record_gaps_end_of_record_iostat() {
    let out = run_prints(
        r#"
program test_io_list_directed_record_gaps
    character(len=10) :: text
    integer :: x
    integer :: ios
    text = '7'
    read(text, *, iostat=ios) x
    read(text, *, iostat=ios) x
    if (ios /= 0) print *, 1
end program test_io_list_directed_record_gaps
"#,
    );

    assert_eq!(out, vec!["1"]);
}
