! vybe-test: fortran/numerics/num_nan_08
! origin: languages/fortran/tests/fortran/test_numerics.rs
program p
real::x
x=0.0/0.0
print *,x
end program p
