! vybe-test: fortran/arrays_shift/eoshift_1d_left_one
! origin: languages/fortran/tests/fortran/test_arrays_shift.rs

program test
    integer :: a(5) = [1, 2, 3, 4, 5]
    integer :: b(5)
    b = eoshift(a, 1)
    print *, b(1)
    print *, b(5)
end program test
