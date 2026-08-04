! vybe-test: fortran/intrinsic_math/log_e
! origin: languages/fortran/tests/fortran/test_intrinsic_math.rs
program t
real :: x
x = log(2.718)
print *, x
end program t
