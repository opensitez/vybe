! vybe-test: fortran/blank_interpretation/blank_interpretation_08
! origin: languages/fortran/tests/fortran/test_blank_interpretation.rs
program p
character(len=6) :: s='ab cd '
print *, trim(s)
end program p
