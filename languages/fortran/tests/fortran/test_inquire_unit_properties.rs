use super::helpers::run_prints;

#[test]
fn test_inquire_unit_properties_opened_flag() {
    let out = run_prints(
        r#"
program test_inquire_unit_properties
    integer :: unit
    logical :: opened
    open(newunit=unit, status='scratch', action='readwrite')
    inquire(unit=unit, opened=opened)
    if (opened) then
        print *, 1
    else
        print *, 0
    end if
    close(unit)
end program test_inquire_unit_properties
"#,
    );

    assert_eq!(out, vec!["1"]);
}
