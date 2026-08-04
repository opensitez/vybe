! vybe-test: fortran/miscellaneous/misc_storage_assoc_02
! origin: languages/fortran/tests/fortran/test_miscellaneous.rs
program p
integer::a(2),b
equivalence(a(1),b)
print *,1
end program p
