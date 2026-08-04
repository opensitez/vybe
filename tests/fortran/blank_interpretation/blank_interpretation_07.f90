! vybe-test: fortran/blank_interpretation/blank_interpretation_07
! origin: languages/fortran/tests/fortran/test_blank_interpretation.rs
program p
character(len=5) :: s='     '
print *, len_trim(s)
end program p
