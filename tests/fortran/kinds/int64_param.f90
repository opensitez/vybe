! vybe-test: fortran/kinds/int64_param
! origin: languages/fortran/tests/fortran/test_kinds.rs

program test
    integer, parameter :: int64 = 8
    integer(kind=int64) :: big = 100000000000_8
    print *, big
end program test
