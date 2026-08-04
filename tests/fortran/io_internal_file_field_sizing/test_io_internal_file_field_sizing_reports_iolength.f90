! vybe-test: fortran/io_internal_file_field_sizing/test_io_internal_file_field_sizing_reports_iolength
! origin: languages/fortran/tests/fortran/test_io_internal_file_field_sizing.rs

program test_io_internal_file_field_sizing
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 1 ]
    integer :: required
    integer :: code
    integer :: value
    value = 256
    inquire(iolength=required) value
    if (required > 0) code = 1
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 1) then
        print *, "FAIL: more than 1 line(s)"
        stop 1
    end if
    if ((code) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", code, "]"
        stop 1
    end if
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 1) then
        print *, "FAIL: more than 1 line(s)"
        stop 1
    end if
    if ((required) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", required, "]"
        stop 1
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program test_io_internal_file_field_sizing
