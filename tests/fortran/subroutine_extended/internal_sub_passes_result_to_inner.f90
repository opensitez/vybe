! vybe-test: fortran/subroutine_extended/internal_sub_passes_result_to_inner
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
call driver(4)
contains
function triple(n) result(r)
integer, intent(in) :: n
integer :: r
r = n * 3
end function triple
subroutine driver(n)
integer, intent(in) :: n
call show(triple(n))
end subroutine driver
subroutine show(v)
integer, intent(in) :: v
if ((v) /= 12) then
    print *, "FAIL: want [12] got [", v, "]"
    stop 1
end if
end subroutine show
end program t
