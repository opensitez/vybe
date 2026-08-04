! vybe-test: fortran/functions/recursive_function
! origin: languages/fortran/tests/fortran/test_functions.rs

program test
    print *, factorial(5)
contains
    recursive function factorial(n) result(res)
        integer, intent(in) :: n
        integer :: res
        if (n <= 1) then
            res = 1
        else
            res = n * factorial(n - 1)
        end if
    end function factorial
end program test
