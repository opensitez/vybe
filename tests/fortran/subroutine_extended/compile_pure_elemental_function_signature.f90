! vybe-test: fortran/subroutine_extended/compile_pure_elemental_function_signature
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs

program t
    integer :: a(2) = [1, 2]
    if ((blend(a, 0)) /= 3) then
    print *, "FAIL: want [3] got [", blend(a, 0), "]"
    stop 1
end if
contains
    elemental function blend(x, bias) result(r)
        integer, intent(in) :: x, bias
        integer :: r
        r = x + bias
end function blend
end program t
