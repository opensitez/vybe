! vybe-test: fortran/intent_optional_extended/optional_middle_arg_present_only
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
if ((tri_opt(1, 5)) /= 6) then
    print *, "FAIL: want [6] got [", tri_opt(1, 5), "]"
    stop 1
end if
contains
integer function tri_opt(a, b, c)
integer, intent(in) :: a
integer, intent(in), optional :: b, c
integer :: s
s = a
if (present(b)) s = s + b
if (present(c)) s = s + c
tri_opt = s
end function tri_opt
end program t
