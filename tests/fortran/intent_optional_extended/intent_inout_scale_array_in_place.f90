! vybe-test: fortran/intent_optional_extended/intent_inout_scale_array_in_place
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 24 ]
integer :: a(3)
a = [2, 4, 6]
call scale_inplace(a, 3)
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((sum(a)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", sum(a), "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
contains
subroutine scale_inplace(v, n)
integer, intent(inout) :: v(n)
integer, intent(in) :: n
integer :: i
do i = 1, n
v(i) = v(i) * 2
end do
end subroutine scale_inplace
end program t
