! vybe-test: fortran/padding_rules/padding_rules_05
! origin: languages/fortran/tests/fortran/test_padding_rules.rs
program p
character(len=6) :: s='hi'
print *, len_trim(s)
end program p
