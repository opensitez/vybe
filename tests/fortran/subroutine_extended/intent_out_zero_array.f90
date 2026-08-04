! vybe-test: fortran/subroutine_extended/intent_out_zero_array
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
integer :: a(3)
call zero_fill(a)
if ((sum(a)) /= 0) then
    print *, "FAIL: want [0] got [", sum(a), "]"
    stop 1
end if
contains
subroutine zero_fill(arr)
integer, intent(out) :: arr(3)
arr = 0
end subroutine zero_fill
end program t
