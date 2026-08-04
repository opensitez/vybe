! vybe-test: fortran/module_use_extended/explicit_public_subroutine_callable
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module greet
implicit none
contains
public :: say_hi
subroutine say_hi()
if ((1) /= 1) then
    print *, "FAIL: want [1] got [", 1, "]"
    stop 1
end if
end subroutine say_hi
end module greet
program t
use greet
call say_hi()
end program t
