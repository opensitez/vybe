! vybe-test: fortran/variables/character_assign
! origin: languages/fortran/tests/fortran/test_variables.rs
program t
character(len=20) :: s
s = "world"
if (trim(s) /= "world") then
    print *, "FAIL: want [world] got [", s, "]"
    stop 1
end if
end program t
