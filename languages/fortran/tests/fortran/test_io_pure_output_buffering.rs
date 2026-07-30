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

#[test]
fn test_io_pure_output_buffering_mixed_stdout_writes() {
    let out = run_prints(
        r#"
program test_io_pure_output_buffering
    write(*, '(I0)') 7
    write(*, '(A)', advance='no') 'x'
    write(*, '(A)') 'y'
end program test_io_pure_output_buffering
"#,
    );

    assert_eq!(out, vec!["7", "xy"]);
}

#[test]
fn test_io_pure_output_buffering_buffered_file_roundtrip() {
    let out = run_prints(
        r#"
program test_io_pure_output_buffering
    integer :: unit
    character(len=10) :: a
    open(newunit=unit, file='pure_buf.txt', status='replace', action='readwrite')
    write(unit, '(I0)') 123
    write(unit, '(A)', advance='no') 'ab'
    write(unit, '(A)') 'cd'
    rewind(unit)
    read(unit, '(A)') a
    close(unit, status='delete')
    print *, trim(a)
end program test_io_pure_output_buffering
"#,
    );

    assert_eq!(out, vec!["123abcd"]);
}
