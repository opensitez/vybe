! vybe-test: fortran/subroutine_extended/recursive_countdown_subroutine
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(3) = [ 3, 2, 1 ]
call count_down(3)
if (vybe_check_i /= 3) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 3"
    stop 1
end if
contains
recursive subroutine count_down(n)
integer, intent(in) :: n
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 3) then
    print *, "FAIL: more than 3 line(s)"
    stop 1
end if
if ((n) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", n, "]"
    stop 1
end if
if (n > 1) call count_down(n - 1)
end subroutine count_down
end program t
