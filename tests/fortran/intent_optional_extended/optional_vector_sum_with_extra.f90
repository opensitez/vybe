! vybe-test: fortran/intent_optional_extended/optional_vector_sum_with_extra
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 16 ]
integer :: v(3)
v = [1, 2, 3]
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((sum_opt(v, 3, 10)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", sum_opt(v, 3, 10), "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
contains
integer function sum_opt(a, n, extra)
integer, intent(in) :: a(n), n
integer, intent(in), optional :: extra
integer :: i
sum_opt = 0
do i = 1, n
sum_opt = sum_opt + a(i)
end do
if (present(extra)) sum_opt = sum_opt + extra
end function sum_opt
end program t
