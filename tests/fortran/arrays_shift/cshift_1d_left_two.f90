! vybe-test: fortran/arrays_shift/cshift_1d_left_two
! origin: languages/fortran/tests/fortran/test_arrays_shift.rs

program test
    integer :: a(5) = [1, 2, 3, 4, 5]
    integer :: b(5)
    b = cshift(a, 2)
    print *, b(1)
end program test
