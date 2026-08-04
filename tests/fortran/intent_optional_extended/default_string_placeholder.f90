! vybe-test: fortran/intent_optional_extended/default_string_placeholder
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
if (trim(trim(name_or('Ada'))) /= "Ada") then
    print *, "FAIL: want [Ada] got [", trim(name_or('Ada')), "]"
    stop 1
end if
if (trim(trim(name_or('Ada', 'Unknown'))) /= "Unknown") then
    print *, "FAIL: want [Unknown] got [", trim(name_or('Ada', 'Unknown')), "]"
    stop 1
end if
contains
character(len=20) function name_or(got, fallback)
character(len=*), intent(in) :: got
character(len=*), intent(in), optional :: fallback
if (present(fallback)) then
name_or = fallback
else
name_or = got
end if
end function name_or
end program t
