! vybe-test: fortran/module_use_extended/program_calls_two_module_procedures
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module duo
implicit none
contains
function inc(n) result(r)
integer, intent(in) :: n
integer :: r
r = n + 1
end function inc
function dec(n) result(r)
integer, intent(in) :: n
integer :: r
r = n - 1
end function dec
end module duo
program t
use duo
if ((inc(dec(6))) /= 6) then
    print *, "FAIL: want [6] got [", inc(dec(6)), "]"
    stop 1
end if
end program t
