! vybe-test: fortran/reduce_intrinsic/reduce_sum
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    integer :: a(5) = [1, 2, 3, 4, 5]
    integer :: total
    total = reduce(a, operator(+))
    print *, total
end program test
