! vybe-test: fortran/variables/character_init
! origin: languages/fortran/tests/fortran/test_variables.rs
program t
character(len=10) :: s = "hello"
if (trim(s) /= "hello") then
    print *, "FAIL: want [hello] got [", s, "]"
    stop 1
end if
end program t
