! vybe-test: fortran/intent_optional_extended/optional_three_tail_args_partial
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
if ((quad_sum(1, 2)) /= 3) then
    print *, "FAIL: want [3] got [", quad_sum(1, 2), "]"
    stop 1
end if
contains
integer function quad_sum(w, x, y, z)
integer, intent(in) :: w
integer, intent(in), optional :: x, y, z
quad_sum = w
if (present(x)) quad_sum = quad_sum + x
if (present(y)) quad_sum = quad_sum + y
if (present(z)) quad_sum = quad_sum + z
end function quad_sum
end program t
