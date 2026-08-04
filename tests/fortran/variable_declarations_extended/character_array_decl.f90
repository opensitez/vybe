! vybe-test: fortran/variable_declarations_extended/character_array_decl
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
character(len=4) :: names(2)
names(1) = "alpha"
names(2) = "beta"
if (trim(trim(names(2))) /= "beta") then
    print *, "FAIL: want [beta] got [", trim(names(2)), "]"
    stop 1
end if
end program t
