! vybe-test: fortran/fortran2018/reduce_with_identity
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    integer :: a(4) = [1, 2, 3, 4]
    integer :: total
    total = reduce(a, operator(+), identity=0)
    print *, total
end program test
