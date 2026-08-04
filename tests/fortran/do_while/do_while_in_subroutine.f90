! vybe-test: fortran/do_while/do_while_in_subroutine
! origin: languages/fortran/tests/fortran/test_do_while.rs

program test
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 55 ]
    integer :: result
    call compute(result)
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 1) then
        print *, "FAIL: more than 1 line(s)"
        stop 1
    end if
    if ((result) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", result, "]"
        stop 1
    end if
contains
    subroutine compute(r)
        integer, intent(out) :: r
        integer :: n = 0
        r = 0
        do while (n < 10)
            n = n + 1
            r = r + n
        end do
    end subroutine compute
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program test
