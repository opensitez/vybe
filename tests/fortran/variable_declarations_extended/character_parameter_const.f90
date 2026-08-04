! vybe-test: fortran/variable_declarations_extended/character_parameter_const
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
character(len=5), parameter :: greeting = "hello"
if (trim(greeting) /= "hello") then
    print *, "FAIL: want [hello] got [", greeting, "]"
    stop 1
end if
end program t
