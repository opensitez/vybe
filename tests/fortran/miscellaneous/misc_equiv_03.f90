! vybe-test: fortran/miscellaneous/misc_equiv_03
! origin: languages/fortran/tests/fortran/test_miscellaneous.rs
program p
integer::a,b
equivalence(a,b)
print *,1
end program p
