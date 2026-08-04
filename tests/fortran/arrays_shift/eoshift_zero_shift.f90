! vybe-test: fortran/arrays_shift/eoshift_zero_shift
! origin: languages/fortran/tests/fortran/test_arrays_shift.rs

program test
    integer :: a(4) = [1, 2, 3, 4]
    integer :: b(4)
    b = eoshift(a, 0)
    print *, b(1)
    print *, b(4)
end program test
