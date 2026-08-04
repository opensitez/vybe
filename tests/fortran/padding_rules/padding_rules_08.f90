! vybe-test: fortran/padding_rules/padding_rules_08
! origin: languages/fortran/tests/fortran/test_padding_rules.rs
program p
character(len=5) :: s='abc'
print *, adjustl(s)
end program p
