! vybe-test: fortran/variables/character_length_truncation_runtime
! origin: languages/fortran/tests/fortran/test_variables.rs
program t
character(len=3) :: s = 'hello'
if (trim(s) /= "hel") then
    print *, "FAIL: want [hel] got [", s, "]"
    stop 1
end if
end program t
