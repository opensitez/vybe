! vybe-test: fortran/padding_rules/padding_rules_04
! origin: languages/fortran/tests/fortran/test_padding_rules.rs
program p
character(len=4) :: s
s='xy'//'z'
print *, s
end program p
