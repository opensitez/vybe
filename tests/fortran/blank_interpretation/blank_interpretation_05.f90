! vybe-test: fortran/blank_interpretation/blank_interpretation_05
! origin: languages/fortran/tests/fortran/test_blank_interpretation.rs
program p
character(len=5) :: s='abc  '
print *, adjustr(s)
end program p
