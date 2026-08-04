! vybe-test: fortran/functions/subroutine_with_args
! origin: languages/fortran/tests/fortran/test_functions.rs

program test
    integer :: a, b, c
    a = 3
    b = 4
    call add_nums(a, b, c)
    if ((c) /= 7) then
    print *, "FAIL: want [7] got [", c, "]"
    stop 1
end if
contains
    subroutine add_nums(x, y, result)
        integer, intent(in) :: x, y
        integer, intent(out) :: result
        result = x + y
    end subroutine add_nums
end program test
