! vybe-test: fortran/module_use_extended/rename_subroutine_to_local_name
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module outmod
implicit none
contains
subroutine show_value(v)
integer, intent(in) :: v
if ((v) /= 19) then
    print *, "FAIL: want [19] got [", v, "]"
    stop 1
end if
end subroutine show_value
end module outmod
program t
use outmod, disp => show_value
call disp(19)
end program t
