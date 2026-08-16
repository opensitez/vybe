! vybe-test: fortran/derived_types/dt_recursive_09
! origin: languages/fortran/tests/fortran/test_derived_types.rs
program t
type::node
integer::x
type(node),pointer::next
end type node
type(node), target :: head
type(node), target :: tail
head%x = 1
tail%x = 2
nullify(tail%next)
head%next => tail
if (head%next%x /= 2) then
    print *, "FAIL: want [2] got [", head%next%x, "]"
    stop 1
end if
if (merge(1, 0, associated(tail%next)) /= 0) then
    print *, "FAIL: want [0] got [", merge(1, 0, associated(tail%next)), "]"
    stop 1
end if
end program t
