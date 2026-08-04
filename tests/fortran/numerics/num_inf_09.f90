! vybe-test: fortran/numerics/num_inf_09
! origin: languages/fortran/tests/fortran/test_numerics.rs
program p
real::x
x=1.0/0.0
print *,x
end program p
