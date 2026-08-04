! vybe-test: fortran/variable_declarations_extended/character_kind_1_len_3
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
character(kind=1, len=3) :: tag = "abc"
if (trim(tag) /= "abc") then
    print *, "FAIL: want [abc] got [", tag, "]"
    stop 1
end if
end program t
