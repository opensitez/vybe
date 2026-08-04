! vybe-test: fortran/internal_io/internal_read_iostat_success_and_fail
! origin: languages/fortran/tests/fortran/test_internal_io.rs

program test
integer :: vybe_check_i = 0
integer :: vybe_check_w(2) = [ 0, 1 ]
    character(len=8) :: ok = ' 12'
    character(len=8) :: bad = 'abc'
    integer :: x
    integer :: ios_ok, ios_bad
    read(ok, *, iostat=ios_ok) x
    read(bad, *, iostat=ios_bad) x
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 2) then
        print *, "FAIL: more than 2 line(s)"
        stop 1
    end if
    if ((ios_ok) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", ios_ok, "]"
        stop 1
    end if
    if (ios_bad /= 0) then
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 2) then
            print *, "FAIL: more than 2 line(s)"
            stop 1
        end if
        if ((1) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", 1, "]"
            stop 1
        end if
    else
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 2) then
            print *, "FAIL: more than 2 line(s)"
            stop 1
        end if
        if ((0) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", 0, "]"
            stop 1
        end if
    end if
if (vybe_check_i /= 2) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 2"
    stop 1
end if
end program test
