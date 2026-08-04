! vybe-test: fortran/modules/module_with_contains
! origin: languages/fortran/tests/fortran/test_modules.rs
module utils
implicit none
contains
function sq(x) result(r)
real, intent(in) :: x
real :: r
r = x * x
end function sq
end module utils
program t
use utils
print *, sq(5.0)
end program t
