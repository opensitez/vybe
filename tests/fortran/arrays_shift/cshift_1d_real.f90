! vybe-test: fortran/arrays_shift/cshift_1d_real
! origin: languages/fortran/tests/fortran/test_arrays_shift.rs

program test
    real :: a(4) = [1.0, 2.0, 3.0, 4.0]
    real :: b(4)
    b = cshift(a, 1)
    print *, b(1)
end program test
