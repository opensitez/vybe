! vybe-test: fortran/arrays_shift/cshift_in_expression
! origin: languages/fortran/tests/fortran/test_arrays_shift.rs

program test
    integer :: a(4) = [1, 2, 3, 4]
    integer :: b(4)
    b = a + cshift(a, 1)
    print *, b(1)
end program test
