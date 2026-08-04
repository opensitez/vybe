! vybe-test: fortran/pointer_alloc_extended/move_alloc_preserves_first_element
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
integer, allocatable :: from(:), to(:)
allocate(from(3))
from = [11, 22, 33]
call move_alloc(from, to)
if ((to(1)) /= 11) then
    print *, "FAIL: want [11] got [", to(1), "]"
    stop 1
end if
if ((size(to)) /= 3) then
    print *, "FAIL: want [3] got [", size(to), "]"
    stop 1
end if
end program t
