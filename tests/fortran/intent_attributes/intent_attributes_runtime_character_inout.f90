! vybe-test: fortran/intent_attributes/intent_attributes_runtime_character_inout
! origin: languages/fortran/tests/fortran/test_intent_attributes.rs

program test_intent_attributes
character(len=8) :: word = "a"
call append_char(word)
if (trim(word) /= "abc") then
    print *, "FAIL: want [abc] got [", word, "]"
    stop 1
end if

contains
subroutine append_char(s)
character(len=*), intent(inout) :: s
s = trim(s) // "bc"
end subroutine append_char
end program test_intent_attributes
