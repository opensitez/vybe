! vybe-test: fortran/fortran2018/reduce_product
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    integer :: a(4) = [1, 2, 3, 4]
    integer :: prod
    prod = reduce(a, operator(*))
    print *, prod
end program test
