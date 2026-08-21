! vybe-test: fortran/reduce_intrinsic/reduce_max
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    integer :: a(5) = [3, 1, 4, 1, 5]
    integer :: m
    m = reduce(a, my_max)
    print *, m
contains
    pure function my_max(x, y) result(r)
        integer, intent(in) :: x, y
        integer :: r
        r = max(x, y)
    end function my_max
end program test
