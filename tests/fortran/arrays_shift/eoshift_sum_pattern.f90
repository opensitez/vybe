! vybe-test: fortran/arrays_shift/eoshift_sum_pattern
! origin: languages/fortran/tests/fortran/test_arrays_shift.rs

program test
    real :: a(5) = [1.0, 2.0, 3.0, 4.0, 5.0]
    real :: forward(5), backward(5), centered(5)
    forward  = eoshift(a,  1, boundary=0.0)
    backward = eoshift(a, -1, boundary=0.0)
    centered = (forward + backward) * 0.5
    print *, centered(3)
end program test
