! vybe-test: fortran/intent_optional_extended/intent_in_parameter_passed_to_len
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
if ((boxed_len('Fortran')) /= 9) then
    print *, "FAIL: want [9] got [", boxed_len('Fortran'), "]"
    stop 1
end if
contains
integer function boxed_len(text)
character(len=*), intent(in) :: text
boxed_len = len_trim(text) + 2
end function boxed_len
end program t
