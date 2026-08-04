! vybe-test: fortran/where_advanced/storage_size_int8
! origin: languages/fortran/tests/fortran/test_where_advanced.rs

program test
    integer(kind=8) :: big = 0_8
    integer(kind=4) :: small = 0_4
    print *, storage_size(big) > storage_size(small)
end program test
