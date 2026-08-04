! vybe-test: fortran/intent_optional_extended/present_count_two_optional_args
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
if ((count_present(1)) /= 0) then
    print *, "FAIL: want [0] got [", count_present(1), "]"
    stop 1
end if
if ((count_present(1, 2)) /= 1) then
    print *, "FAIL: want [1] got [", count_present(1, 2), "]"
    stop 1
end if
if ((count_present(1, 2, 3)) /= 2) then
    print *, "FAIL: want [2] got [", count_present(1, 2, 3), "]"
    stop 1
end if
contains
integer function count_present(a, b, c)
integer, intent(in) :: a
integer, intent(in), optional :: b, c
count_present = 0
if (present(b)) count_present = count_present + 1
if (present(c)) count_present = count_present + 1
end function count_present
end program t
