! vybe-test: fortran/pointer_alloc_extended/pointer_derived_next_link_value
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
type :: Link
integer :: payload
type(Link), pointer :: nxt => null()
end type Link
type(Link), target :: head, tail
head%payload = 1
tail%payload = 2
head%nxt => tail
if ((head%nxt%payload) /= 2) then
    print *, "FAIL: want [2] got [", head%nxt%payload, "]"
    stop 1
end if
end program t
