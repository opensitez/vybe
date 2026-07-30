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

#[test]
fn test_inquire_unit_properties_access_and_form() {
    let out = run_prints(
        r#"
program test_inquire_unit_properties
    integer :: unit
    character(len=20) :: acc, frm
    open(newunit=unit, status='scratch')
    inquire(unit=unit, access=acc, form=frm)
    close(unit)
    print *, trim(acc)
    print *, trim(frm)
end program test_inquire_unit_properties
"#,
    );

    assert_eq!(out, vec!["SEQUENTIAL", "FORMATTED"]);
}

#[test]
fn test_inquire_unit_properties_action_mode() {
    let out = run_prints(
        r#"
program test_inquire_unit_properties
    integer :: unit
    character(len=20) :: action_mode
    open(newunit=unit, status='scratch', action='readwrite')
    inquire(unit=unit, action=action_mode)
    close(unit)
    print *, trim(action_mode)
end program test_inquire_unit_properties
"#,
    );

    assert_eq!(out, vec!["READWRITE"]);
}

#[test]
fn test_inquire_unit_properties_closed_unit_iostat() {
    let out = run_prints(
        r#"
program test_inquire_unit_properties
    integer :: unit
    integer :: ios
    logical :: opened
    open(newunit=unit, status='scratch')
    close(unit)
    inquire(unit=unit, iostat=ios, opened=opened)
    if (opened) then
        print *, 1
    else
        print *, 0
    end if
    print *, ios
end program test_inquire_unit_properties
"#,
    );

    assert_eq!(out, vec!["0", "0"]);
}
