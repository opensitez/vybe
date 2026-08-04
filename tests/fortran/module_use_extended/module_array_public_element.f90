! vybe-test: fortran/module_use_extended/module_array_public_element
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module arrmod
implicit none
integer :: data(3) = [4, 5, 6]
end module arrmod
program t
use arrmod
if ((data(2)) /= 5) then
    print *, "FAIL: want [5] got [", data(2), "]"
    stop 1
end if
end program t
