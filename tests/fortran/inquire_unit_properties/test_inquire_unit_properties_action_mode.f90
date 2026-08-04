! vybe-test: fortran/inquire_unit_properties/test_inquire_unit_properties_action_mode
! origin: languages/fortran/tests/fortran/test_inquire_unit_properties.rs

program test_inquire_unit_properties
    integer :: unit
    character(len=20) :: action_mode
    open(newunit=unit, status='scratch', action='readwrite')
    inquire(unit=unit, action=action_mode)
    close(unit)
    if (trim(trim(action_mode)) /= "READWRITE") then
    print *, "FAIL: want [READWRITE] got [", trim(action_mode), "]"
    stop 1
end if
end program test_inquire_unit_properties
