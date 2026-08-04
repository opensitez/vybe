! vybe-test: fortran/arrays_shift/cshift_2d_negative_dim2
! origin: languages/fortran/tests/fortran/test_arrays_shift.rs

program test
    integer :: m(2,4) = reshape([1,2,3,4,5,6,7,8],[2,4])
    integer :: n(2,4)
    n = cshift(m, -1, dim=2)
    print *, n(1,1)
end program test
