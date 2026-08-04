! vybe-test: fortran/module_use_extended/use_only_logical_flag
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module flags
implicit none
logical :: active = .true.
logical :: debug = .false.
end module flags
program t
use flags, only: active
if ((active) /= 1) then
    print *, "FAIL: want [1] got [", active, "]"
    stop 1
end if
end program t
