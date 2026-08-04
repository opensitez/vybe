! vybe-test: fortran/padding_rules/padding_rules_06
! origin: languages/fortran/tests/fortran/test_padding_rules.rs
program p
character(len=4) :: a(2)
a(1)='a'
a(2)='bc'
print *, a
end program p
