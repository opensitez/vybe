! vybe-test: fortran/inquire_unit_properties/test_inquire_unit_properties_access_and_form
! origin: languages/fortran/tests/fortran/test_inquire_unit_properties.rs

program test_inquire_unit_properties
    integer :: unit
    character(len=20) :: acc, frm
    open(newunit=unit, status='scratch')
    inquire(unit=unit, access=acc, form=frm)
    close(unit)
    if (trim(trim(acc)) /= "SEQUENTIAL") then
    print *, "FAIL: want [SEQUENTIAL] got [", trim(acc), "]"
    stop 1
end if
    if (trim(trim(frm)) /= "FORMATTED") then
    print *, "FAIL: want [FORMATTED] got [", trim(frm), "]"
    stop 1
end if
end program test_inquire_unit_properties
