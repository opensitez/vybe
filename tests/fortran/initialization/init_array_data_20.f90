! vybe-test: fortran/initialization/init_array_data_20
! origin: languages/fortran/tests/fortran/test_initialization.rs
program p
integer::a(3)
data a/1,2,3/
print *,a
end program p
