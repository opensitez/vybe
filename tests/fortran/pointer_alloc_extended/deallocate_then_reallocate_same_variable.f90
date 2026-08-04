! vybe-test: fortran/pointer_alloc_extended/deallocate_then_reallocate_same_variable
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
integer, allocatable :: v(:)
allocate(v(2))
v = [5, 6]
deallocate(v)
allocate(v(4))
v = [(i, i = 1, 4)]
if ((sum(v)) /= 10) then
    print *, "FAIL: want [10] got [", sum(v), "]"
    stop 1
end if
deallocate(v)
end program t
