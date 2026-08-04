! vybe-test: fortran/module_use_extended/rename_function_to_short_alias
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module mathfn
implicit none
contains
function cube(n) result(r)
integer, intent(in) :: n
integer :: r
r = n * n * n
end function cube
end module mathfn
program t
use mathfn, cb => cube
if ((cb(3)) /= 27) then
    print *, "FAIL: want [27] got [", cb(3), "]"
    stop 1
end if
end program t
