! vybe-test: fortran/functions/function_name_result_runtime
! origin: languages/fortran/tests/fortran/test_functions.rs

program test
    if ((cube(3)) /= 27) then
    print *, "FAIL: want [27] got [", cube(3), "]"
    stop 1
end if
contains
    integer function cube(n)
        integer, intent(in) :: n
        cube = n * n * n
    end function cube
end program test
