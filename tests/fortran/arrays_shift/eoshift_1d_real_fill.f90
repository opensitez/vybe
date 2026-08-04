! vybe-test: fortran/arrays_shift/eoshift_1d_real_fill
! origin: languages/fortran/tests/fortran/test_arrays_shift.rs

program test
    real :: a(4) = [1.0, 2.0, 3.0, 4.0]
    real :: b(4)
    b = eoshift(a, 1, boundary=0.0)
    print *, b(4)
end program test
