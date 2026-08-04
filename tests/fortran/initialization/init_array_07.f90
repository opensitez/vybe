! vybe-test: fortran/initialization/init_array_07
! origin: languages/fortran/tests/fortran/test_initialization.rs
program p
integer::a(3)=[1,2,3]
print *,a
end program p
