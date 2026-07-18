use super::helpers::run_prints;

#[test]
fn test_file_positioning_and_backspace_moves_cursor() {
    let out = run_prints(
        r#"
program test_file_positioning_and_backspace
    integer :: unit
    integer :: v1
    integer :: v2
    open(newunit=unit, status='scratch', action='readwrite')
    write(unit, '(I0)') 12
    write(unit, '(I0)') 34
    backspace(unit)
    read(unit, '(I0)') v1
    rewind(unit)
    read(unit, '(I0)') v2
    print *, v1
    print *, v2
    close(unit)
end program test_file_positioning_and_backspace
"#,
    );

    assert_eq!(out, vec!["34", "12"]);
}
