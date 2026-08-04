! vybe-test: fortran/variable_declarations_extended/parameter_character_literal
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
character(len=3), parameter :: tag = "vyb"
if (trim(tag) /= "vyb") then
    print *, "FAIL: want [vyb] got [", tag, "]"
    stop 1
end if
end program t
