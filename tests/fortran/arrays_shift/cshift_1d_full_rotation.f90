! vybe-test: fortran/arrays_shift/cshift_1d_full_rotation
! origin: languages/fortran/tests/fortran/test_arrays_shift.rs

program test
    integer :: a(4) = [1, 2, 3, 4]
    integer :: b(4)
    b = cshift(a, 4)
    print *, b(1)
end program test
