! vybe-test: fortran/inquire_file_property_matrix/test_inquire_file_property_matrix_file_exists_after_create
! origin: languages/fortran/tests/fortran/test_inquire_file_property_matrix.rs

program test_inquire_file_property_matrix
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 1 ]
    logical :: exists
    character(len=256) :: path
    path = 'tmp_inquire_exists_probe.txt'
    open(unit=99, file=trim(path), status='replace')
    close(99)
    inquire(file=trim(path), exist=exists)
    if (exists) then
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
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program test_inquire_file_property_matrix
