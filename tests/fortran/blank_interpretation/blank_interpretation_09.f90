! vybe-test: fortran/blank_interpretation/blank_interpretation_09
! origin: languages/fortran/tests/fortran/test_blank_interpretation.rs
program p
character(len=6) :: s='ab cd '
print *, s(3:4)
end program p
