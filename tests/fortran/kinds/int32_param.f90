! vybe-test: fortran/kinds/int32_param
! origin: languages/fortran/tests/fortran/test_kinds.rs

program test
    integer, parameter :: int32 = 4
    integer(kind=int32) :: x = 2147483647
    print *, x
end program test
