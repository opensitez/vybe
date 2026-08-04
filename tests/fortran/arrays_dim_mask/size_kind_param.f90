! vybe-test: fortran/arrays_dim_mask/size_kind_param
! origin: languages/fortran/tests/fortran/test_arrays_dim_mask.rs

program test
    use iso_fortran_env
    integer :: a(100)
    integer(int64) :: n
    n = size(a, kind=int64)
    print *, n
end program test
