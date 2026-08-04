! vybe-test: fortran/module_use_extended/use_only_subroutine_from_module
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module io_mod
implicit none
contains
subroutine emit(v)
integer, intent(in) :: v
if ((v) /= 77) then
    print *, "FAIL: want [77] got [", v, "]"
    stop 1
end if
end subroutine emit
end module io_mod
program t
use io_mod, only: emit
call emit(77)
end program t
