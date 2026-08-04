! vybe-test: fortran/pure_elemental/optional_present_check_runtime
! origin: languages/fortran/tests/fortran/test_pure_elemental.rs

program test
integer :: vybe_check_i = 0
integer :: vybe_check_w(2) = [ 3, 7 ]
    call maybe(3)
    call maybe(3, 4)
contains
    subroutine maybe(n, extra)
        integer, intent(in) :: n
        integer, intent(in), optional :: extra
        integer :: total
        total = n
        if (present(extra)) total = total + extra
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 2) then
            print *, "FAIL: more than 2 line(s)"
            stop 1
        end if
        if ((total) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", total, "]"
            stop 1
        end if
    end subroutine maybe
if (vybe_check_i /= 2) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 2"
    stop 1
end if
end program test
