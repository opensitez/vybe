! vybe-test: fortran/subroutine_extended/intent_out_pair_scalars
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
integer :: lo, hi
call bounds(lo, hi)
if ((lo) /= 3) then
    print *, "FAIL: want [3] got [", lo, "]"
    stop 1
end if
if ((hi) /= 11) then
    print *, "FAIL: want [11] got [", hi, "]"
    stop 1
end if
contains
subroutine bounds(a, b)
integer, intent(out) :: a, b
a = 3
b = 11
end subroutine bounds
end program t
