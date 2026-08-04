! vybe-test: fortran/arrays_shift/cshift_2d_dim1
! origin: languages/fortran/tests/fortran/test_arrays_shift.rs

program test
    integer :: m(3,4) = reshape([(i, i=1,12)],[3,4])
    integer :: n(3,4)
    n = cshift(m, 1, dim=1)
    print *, n(1,1)
end program test
