! vybe-test: fortran/procedure_attributes/module_procedure_vs_external_same_shape
! origin: languages/fortran/tests/fortran/test_procedure_attributes.rs
module attr_mod
implicit none
interface add_pair
module procedure mod_add
end interface
contains
function mod_add(a, b) result(r)
integer, intent(in) :: a, b
integer :: r
r = a + b
end function mod_add
end module attr_mod
function ext_add(a, b) result(r)
integer, intent(in) :: a, b
integer :: r
r = a * b
end function ext_add
program t
use attr_mod
if ((add_pair(3, 4)) /= 7) then
    print *, "FAIL: want [7] got [", add_pair(3, 4), "]"
    stop 1
end if
if ((ext_add(3, 4)) /= 12) then
    print *, "FAIL: want [12] got [", ext_add(3, 4), "]"
    stop 1
end if
end program t
