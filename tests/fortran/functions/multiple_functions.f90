! vybe-test: fortran/functions/multiple_functions
! origin: languages/fortran/tests/fortran/test_functions.rs

program test
    print *, add(3, 4)
    print *, multiply(3, 4)
contains
    function add(a, b) result(res)
        integer, intent(in) :: a, b
        integer :: res
        res = a + b
    end function add
    function multiply(a, b) result(res)
        integer, intent(in) :: a, b
        integer :: res
        res = a * b
    end function multiply
end program test
