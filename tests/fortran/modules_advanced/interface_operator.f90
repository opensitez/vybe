! vybe-test: fortran/modules_advanced/interface_operator
! origin: languages/fortran/tests/fortran/test_modules_advanced.rs

module vectors
    implicit none
    type :: Vec
        real :: x, y
    end type Vec
    interface operator(+)
        module procedure add_vecs
    end interface
contains
    function add_vecs(a, b) result(c)
        type(Vec), intent(in) :: a, b
        type(Vec) :: c
        c%x = a%x + b%x
        c%y = a%y + b%y
    end function add_vecs
end module vectors

program test
    use vectors
    type(Vec) :: v1, v2, v3
    v1 = Vec(1.0, 2.0)
    v2 = Vec(3.0, 4.0)
    v3 = v1 + v2
    if ((v3%x) /= 4) then
    print *, "FAIL: want [4] got [", v3%x, "]"
    stop 1
end if
end program test
