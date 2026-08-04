! vybe-test: fortran/initialization/init_data_05
! origin: languages/fortran/tests/fortran/test_initialization.rs
program p
integer::x
data x/1/
print *,x
end program p
