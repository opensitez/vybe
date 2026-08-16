! vybe-test: fortran/legacy_data_extended/save_persists_across_calls_runtime
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs

program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(3) = [ 1, 2, 3 ]
    call tick()
    call tick()
    call tick()
if (vybe_check_i /= 3) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 3"
    stop 1
end if
contains
    subroutine tick()
        integer, save :: n = 0
        n = n + 1
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 3) then
            print *, "FAIL: more than 3 line(s)"
            stop 1
        end if
        if ((n) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", n, "]"
            stop 1
        end if
    end subroutine tick
end program t
