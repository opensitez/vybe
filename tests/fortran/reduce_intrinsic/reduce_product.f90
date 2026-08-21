! vybe-test: fortran/reduce_intrinsic/reduce_product
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

! `operator(+)` is a GENERIC-SPEC — legal in INTERFACE/USE/generic bindings,
! NOT as an actual argument. gfortran: "Syntax error in expression". F2018
! REDUCE takes a PURE FUNCTION of two arguments, so that is what it gets.
program test
    integer :: a(4) = [1, 2, 3, 4]
    integer :: prod
    prod = reduce(a, vy_mul)
    print *, prod
contains
    pure function vy_mul(x, y) result(r)
        integer, intent(in) :: x, y
        integer :: r
        r = x * y
    end function vy_mul
end program test
