! vybe-test: fortran/functions/function_with_return_type
! origin: languages/fortran/tests/fortran/test_functions.rs

program test
    print *, cube(3)
contains
    integer function cube(n)
        integer, intent(in) :: n
        cube = n * n * n
    end function cube
end program test
