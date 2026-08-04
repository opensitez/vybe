! vybe-test: fortran/intent_optional_extended/optional_last_of_four_present
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
if ((pack4(1, 2, 3, 4)) /= 10) then
    print *, "FAIL: want [10] got [", pack4(1, 2, 3, 4), "]"
    stop 1
end if
contains
integer function pack4(a, b, c, d)
integer, intent(in) :: a, b, c
integer, intent(in), optional :: d
pack4 = a + b + c
if (present(d)) pack4 = pack4 + d
end function pack4
end program t
