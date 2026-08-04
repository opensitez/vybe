! vybe-test: fortran/interface_operator_extended/module_interface_subroutine_generic
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gsub
implicit none
interface run
module procedure run_int, run_real
end interface
contains
subroutine run_int(v)
integer, intent(in) :: v
if ((v * 2) /= 10) then
    print *, "FAIL: want [10] got [", v * 2, "]"
    stop 1
end if
end subroutine run_int
subroutine run_real(v)
real, intent(in) :: v
if ((int(v * 2.0)) /= 5) then
    print *, "FAIL: want [5] got [", int(v * 2.0), "]"
    stop 1
end if
end subroutine run_real
end module gsub
program t
use gsub
call run(5)
call run(2.5)
end program t
