! vybe-test: fortran/initialization/init_equivalence_15
! origin: languages/fortran/tests/fortran/test_initialization.rs
program p
integer::a,b
equivalence(a,b)
a=1
print *,b
end program p
