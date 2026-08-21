! vybe-test: fortran/type_bound_procedures/tbp_nopass_scale_helper
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs
program t
type :: Scale
contains
procedure(scale_by), nopass :: apply
end type Scale
type(Scale) :: s
if ((s%apply(6, 2)) /= 12) then
    print *, "FAIL: want [12] got [", s%apply(6, 2), "]"
    stop 1
end if
contains
integer function scale_by(n, k) result(r)
integer, intent(in) :: n, k
r = n * k
end function scale_by
end program t
