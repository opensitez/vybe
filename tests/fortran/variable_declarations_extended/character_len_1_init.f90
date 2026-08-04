! vybe-test: fortran/variable_declarations_extended/character_len_1_init
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
character(len=1) :: c = "Z"
if (trim(c) /= "Z") then
    print *, "FAIL: want [Z] got [", c, "]"
    stop 1
end if
end program t
