! vybe-test: fortran/interfaces/if_module_sub_32
! origin: languages/fortran/tests/fortran/test_interfaces.rs
module m
integer :: seen = 0
interface
module subroutine s(x)
integer :: x
end subroutine s
end interface
end module m
submodule (m) msub
contains
module subroutine s(x)
integer :: x
seen = x
end subroutine s
end submodule msub
program t
use m
call s(5)
if (seen /= 5) then
    print *, "FAIL: want [5] got [", seen, "]"
    stop 1
end if
end program t
