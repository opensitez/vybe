! vybe-test: fortran/modules_advanced/operator_plus_type
! origin: languages/fortran/tests/fortran/test_modules_advanced.rs

module complex_mod
    implicit none
    type :: MyComplex
        real :: re, im
    end type MyComplex
    interface operator(+)
        module procedure add_complex
    end interface
contains
    function add_complex(a, b) result(c)
        type(MyComplex), intent(in) :: a, b
        type(MyComplex) :: c
        c%re = a%re + b%re
        c%im = a%im + b%im
    end function add_complex
end module complex_mod

program test
    use complex_mod
    type(MyComplex) :: a, b, c
    a = MyComplex(1.0, 2.0)
    b = MyComplex(3.0, 4.0)
    c = a + b
    if ((c%re) /= 4) then
    print *, "FAIL: want [4] got [", c%re, "]"
    stop 1
end if
end program test
