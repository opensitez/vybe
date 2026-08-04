! vybe-test: fortran/intent_optional_extended/optional_two_first_present
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
if ((combine3(4, 2)) /= 6) then
    print *, "FAIL: want [6] got [", combine3(4, 2), "]"
    stop 1
end if
contains
integer function combine3(a, b, c)
integer, intent(in) :: a
integer, intent(in), optional :: b, c
combine3 = a
if (present(b)) combine3 = combine3 + b
if (present(c)) combine3 = combine3 + c
end function combine3
end program t
