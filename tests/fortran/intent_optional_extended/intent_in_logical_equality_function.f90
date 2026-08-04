! vybe-test: fortran/intent_optional_extended/intent_in_logical_equality_function
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
if ((same_flag(.true., .true.)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", same_flag(.true., .true.), "]"
    stop 1
end if
if ((same_flag(.true., .false.)) .neqv. .false.) then
    print *, "FAIL: want [false] got [", same_flag(.true., .false.), "]"
    stop 1
end if
contains
logical function same_flag(a, b)
logical, intent(in) :: a, b
same_flag = a .eqv. b
end function same_flag
end program t
