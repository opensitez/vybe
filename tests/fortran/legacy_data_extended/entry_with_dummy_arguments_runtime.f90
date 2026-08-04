! vybe-test: fortran/legacy_data_extended/entry_with_dummy_arguments_runtime
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs

program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 5 ]
    call master(5)
contains
    subroutine master(x)
        integer, intent(in) :: x
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if ((x) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", x, "]"
            stop 1
        end if
        return
    entry slave(y)
        integer :: y
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if ((y + 1) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", y + 1, "]"
            stop 1
        end if
    end subroutine master
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
