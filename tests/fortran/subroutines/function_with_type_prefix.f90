! vybe-test: fortran/subroutines/function_with_type_prefix
! origin: languages/fortran/tests/fortran/test_subroutines.rs
program t
if ((add(3, 4)) /= 7) then
    print *, "FAIL: want [7] got [", add(3, 4), "]"
    stop 1
end if
contains
integer function add(a, b)
integer, intent(in) :: a, b
add = a + b
end function add
end program t
