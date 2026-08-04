! vybe-test: fortran/submodule_extended/submodule_generic_iface_present_program_sums_array
! origin: languages/fortran/tests/fortran/test_submodule_extended.rs
module generic_iface
implicit none
interface norm
module function norm_real(a) result(r)
real, intent(in) :: a(:)
real :: r
end function norm_real
end interface norm
end module generic_iface
submodule (generic_iface) generic_impl
contains
module function norm_real(a) result(r)
real, intent(in) :: a(:)
real :: r
r = sqrt(sum(a**2))
end function norm_real
end submodule generic_impl
program t
use generic_iface
real :: v(3) = [3.0, 4.0, 0.0]
if ((int(sum(v))) /= 7) then
    print *, "FAIL: want [7] got [", int(sum(v)), "]"
    stop 1
end if
end program t
