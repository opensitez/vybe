! vybe-test: fortran/arrays_shift/eoshift_2d_with_boundary
! origin: languages/fortran/tests/fortran/test_arrays_shift.rs

program test
    integer :: m(2,4) = reshape([1,2,3,4,5,6,7,8],[2,4])
    integer :: n(2,4)
    n = eoshift(m, 2, boundary=-99, dim=2)
    print *, n(1,3)
    print *, n(1,4)
end program test
