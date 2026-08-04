! vybe-test: fortran/module_use_extended/rename_integer_variable_alias
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module counters
implicit none
integer :: tally = 8
integer :: spare = 0
end module counters
program t
use counters, total => tally
if ((total) /= 8) then
    print *, "FAIL: want [8] got [", total, "]"
    stop 1
end if
end program t
