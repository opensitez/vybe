! vybe-test: fortran/fortran2018/reduce_dim
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    integer :: m(3,3) = reshape([1,2,3,4,5,6,7,8,9],[3,3])
    integer :: row_sums(3)
    row_sums = reduce(m, operator(+), dim=2)
    print *, row_sums(1)
end program test
