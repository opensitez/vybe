! vybe-test: fortran/variable_declarations_extended/init_character_from_parameter
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
character(len=6), parameter :: base = "planet"
character(len=6) :: word = base
if (trim(trim(word)) /= "planet") then
    print *, "FAIL: want [planet] got [", trim(word), "]"
    stop 1
end if
end program t
