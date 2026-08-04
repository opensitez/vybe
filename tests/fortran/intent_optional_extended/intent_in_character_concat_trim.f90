! vybe-test: fortran/intent_optional_extended/intent_in_character_concat_trim
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
if (trim(trim(merge_names('Ada', 'Lovelace'))) /= "Ada Lovelace") then
    print *, "FAIL: want [Ada Lovelace] got [", trim(merge_names('Ada', 'Lovelace')), "]"
    stop 1
end if
contains
character(len=20) function merge_names(first, last)
character(len=*), intent(in) :: first, last
merge_names = trim(first) // ' ' // trim(last)
end function merge_names
end program t
