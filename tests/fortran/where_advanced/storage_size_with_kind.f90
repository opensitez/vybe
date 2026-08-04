! vybe-test: fortran/where_advanced/storage_size_with_kind
! origin: languages/fortran/tests/fortran/test_where_advanced.rs

program test
    use iso_fortran_env
    integer :: x = 0
    integer(int64) :: n
    n = storage_size(x, kind=int64)
    print *, n
end program test
