use super::helpers::run_prints;

#[test]
fn test_file_rewind_and_endfile_behavior() {
    let out = run_prints(
        r#"
program test_file_rewind_and_endfile_behavior
    integer :: unit
    integer :: first
    integer :: second
    integer :: code

    open(newunit=unit, status='scratch', action='readwrite')
    write(unit, '(I0)') 7
    write(unit, '(I0)') 9
    rewind(unit)
    read(unit, '(I0)') first
    endfile(unit)
    read(unit, '(I0)', iostat=code) second

    print *, first
    print *, code
    close(unit)
end program test_file_rewind_and_endfile_behavior
"#,
    );

    assert_eq!(out, vec!["7", "-1"]);
}
