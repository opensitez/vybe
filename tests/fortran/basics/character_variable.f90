! vybe-test: fortran/basics/character_variable
! origin: languages/fortran/tests/fortran/test_basics.rs

program test
    character(len=20) :: name
    name = "Fortran"
    if (trim(name) /= "Fortran") then
    print *, "FAIL: want [Fortran] got [", name, "]"
    stop 1
end if
end program test
