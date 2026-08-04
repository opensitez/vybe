! vybe-test: fortran/intent_optional_extended/present_on_character_optional
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
call tag_len('abc')
call tag_len('abc', 'xyz')
contains
subroutine tag_len(a, b)
character(len=*), intent(in) :: a
character(len=*), intent(in), optional :: b
if (present(b)) then
if ((len_trim(a) + len_trim(b)) /= 3) then
    print *, "FAIL: want [3] got [", len_trim(a) + len_trim(b), "]"
    stop 1
end if
else
if ((len_trim(a)) /= 6) then
    print *, "FAIL: want [6] got [", len_trim(a), "]"
    stop 1
end if
end if
end subroutine tag_len
end program t
