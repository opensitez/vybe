! vybe-test: fortran/intent_optional_extended/intent_out_character_buffer_fill
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
character(len=5) :: tag
call fill_tag(tag)
if (trim(trim(tag)) /= "hello") then
    print *, "FAIL: want [hello] got [", trim(tag), "]"
    stop 1
end if
contains
subroutine fill_tag(s)
character(len=5), intent(out) :: s
s = 'hello'
end subroutine fill_tag
end program t
