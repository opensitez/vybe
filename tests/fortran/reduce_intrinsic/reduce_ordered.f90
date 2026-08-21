! vybe-test: fortran/reduce_intrinsic/reduce_ordered
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    integer :: a(5) = [5, 4, 3, 2, 1]
    integer :: r
    r = reduce(a, operator(+), ordered=.true.)
    print *, r
end program test
