! vybe-test: fortran/inquire_unit_properties/test_inquire_unit_properties_opened_flag
! origin: languages/fortran/tests/fortran/test_inquire_unit_properties.rs

program test_inquire_unit_properties
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 1 ]
    integer :: unit
    logical :: opened
    open(newunit=unit, status='scratch', action='readwrite')
    inquire(unit=unit, opened=opened)
    if (opened) then
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if ((1) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", 1, "]"
            stop 1
        end if
    else
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if ((0) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", 0, "]"
            stop 1
        end if
    end if
    close(unit)
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program test_inquire_unit_properties
