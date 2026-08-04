! vybe-test: fortran/subroutine_extended/intent_in_array_minimum
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 3 ]
integer :: v(4)
v = [8, 3, 11, 5]
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((arr_min(v, 4)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", arr_min(v, 4), "]"
    stop 1
end if
contains
function arr_min(a, n) result(m)
integer, intent(in) :: a(n), n
integer :: m, i
m = a(1)
do i = 2, n
if (a(i) < m) m = a(i)
end do
end function arr_min
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
