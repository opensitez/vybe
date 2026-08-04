! vybe-test: fortran/pointer_alloc_extended/alloc_int_five_element_sum
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
integer, allocatable :: v(:)
allocate(v(5))
v = [(i, i = 1, 5)]
if ((sum(v)) /= 15) then
    print *, "FAIL: want [15] got [", sum(v), "]"
    stop 1
end if
deallocate(v)
end program t
