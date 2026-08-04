! vybe-test: fortran/intent_optional_extended/intent_inout_character_first_char
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
character(len=4) :: word
word = 'test'
call upper_first(word)
if (trim(trim(word)) /= "Test") then
    print *, "FAIL: want [Test] got [", trim(word), "]"
    stop 1
end if
contains
subroutine upper_first(s)
character(len=4), intent(inout) :: s
if (s(1:1) == 't') s(1:1) = 'T'
end subroutine upper_first
end program t
