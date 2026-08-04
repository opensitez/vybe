! vybe-test: fortran/arrays_dim_mask/shape_with_kind
! origin: languages/fortran/tests/fortran/test_arrays_dim_mask.rs

program test
    use iso_fortran_env
    real :: m(3,4)
    integer(int64), allocatable :: sh(:)
    sh = shape(m, kind=int64)
    print *, sh(1), sh(2)
end program test
