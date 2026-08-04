! vybe-test: fortran/arrays_shift/cshift_1d_right_two
! origin: languages/fortran/tests/fortran/test_arrays_shift.rs

program test
    integer :: a(6) = [1, 2, 3, 4, 5, 6]
    integer :: b(6)
    b = cshift(a, -2)
    print *, b(1)
    print *, b(2)
end program test
