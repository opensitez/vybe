! vybe-test: fortran/padding_rules/padding_rules_10
! origin: languages/fortran/tests/fortran/test_padding_rules.rs
program p
character(len=8) :: s
s = repeat('x',3)
print *, s
end program p
