! vybe-test: fortran/subroutine_extended/elemental_signature_for_arrays_runtime
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs

program t
    integer :: a(2) = [1, 2]
    if ((sum(blend(a, 0))) /= 3) then
    print *, "FAIL: want [3] got [", sum(blend(a, 0)), "]"
    stop 1
end if
    if ((sum(blend(a, 4))) /= 11) then
    print *, "FAIL: want [11] got [", sum(blend(a, 4)), "]"
    stop 1
end if
contains
    elemental function blend(x, bias) result(r)
        integer, intent(in) :: x, bias
        integer :: r
        r = x + bias
    end function blend
end program t
