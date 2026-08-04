! vybe-test: fortran/blank_interpretation/blank_interpretation_04
! origin: languages/fortran/tests/fortran/test_blank_interpretation.rs
program p
character(len=5) :: s='  abc'
print *, adjustl(s)
end program p
