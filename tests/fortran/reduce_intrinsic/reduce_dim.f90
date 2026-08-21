! vybe-test: fortran/reduce_intrinsic/reduce_dim
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

! `operator(+)` is a GENERIC-SPEC — legal in INTERFACE/USE/generic bindings,
! NOT as an actual argument. gfortran: "Syntax error in expression". F2018
! REDUCE takes a PURE FUNCTION of two arguments, so that is what it gets.
program test
    integer :: m(3,3) = reshape([1,2,3,4,5,6,7,8,9],[3,3])
    integer :: row_sums(3)
    row_sums = reduce(m, vy_add, dim=2)
    print *, row_sums(1)
contains
    pure function vy_add(x, y) result(r)
        integer, intent(in) :: x, y
        integer :: r
        r = x + y
    end function vy_add
end program test
