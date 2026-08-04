! vybe-test: fortran/intent_optional_extended/intent_out_substring_target
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
character(len=4) :: code
call code_out(code)
if (trim(trim(code)) /= "ABCD") then
    print *, "FAIL: want [ABCD] got [", trim(code), "]"
    stop 1
end if
contains
subroutine code_out(c)
character(len=4), intent(out) :: c
c = 'ABCD'
end subroutine code_out
end program t
