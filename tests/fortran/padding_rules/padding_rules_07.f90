! vybe-test: fortran/padding_rules/padding_rules_07
! origin: languages/fortran/tests/fortran/test_padding_rules.rs
program p
character(len=5) :: s='abc'
print *, trim(s)
end program p
