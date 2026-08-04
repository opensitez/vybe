! vybe-test: fortran/module_use_extended/module_variable_updated_by_procedure
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module tallymod
implicit none
integer :: hits = 0
contains
subroutine register_hit()
hits = hits + 1
end subroutine register_hit
end module tallymod
program t
use tallymod
call register_hit()
call register_hit()
if ((hits) /= 2) then
    print *, "FAIL: want [2] got [", hits, "]"
    stop 1
end if
end program t
