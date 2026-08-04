! vybe-test: fortran/module_use_extended/program_calls_module_subroutine
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module actions
implicit none
contains
subroutine bump(n)
integer, intent(inout) :: n
n = n + 1
end subroutine bump
end module actions
program t
use actions
integer :: x = 4
call bump(x)
if ((x) /= 5) then
    print *, "FAIL: want [5] got [", x, "]"
    stop 1
end if
end program t
