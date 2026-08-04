! vybe-test: fortran/variable_declarations_extended/character_len_8_trim_print
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
character(len=8) :: s = "fortran"
if (trim(trim(s)) /= "fortran") then
    print *, "FAIL: want [fortran] got [", trim(s), "]"
    stop 1
end if
end program t
