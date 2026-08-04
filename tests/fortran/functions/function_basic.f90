! vybe-test: fortran/functions/function_basic
! origin: languages/fortran/tests/fortran/test_functions.rs

program test
    integer :: result
    result = square(5)
    print *, result
contains
    function square(x) result(res)
        integer, intent(in) :: x
        integer :: res
        res = x * x
    end function square
end program test
