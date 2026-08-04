! vybe-test: fortran/associate_construct_extended/associate_module_variable
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
module amod
integer :: shared = 25
contains
subroutine peek()
associate (s => shared)
if ((s) /= 25) then
    print *, "FAIL: want [25] got [", s, "]"
    stop 1
end if
end associate
end subroutine peek
end module amod
program t
use amod
call peek()
end program t
