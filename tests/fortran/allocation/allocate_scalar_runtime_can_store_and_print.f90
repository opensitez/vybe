! vybe-test: fortran/allocation/allocate_scalar_runtime_can_store_and_print
! origin: languages/fortran/tests/fortran/test_allocation.rs
program t
integer, allocatable :: x
allocate(x)
x = 17
if ((x) /= 17) then
    print *, "FAIL: want [17] got [", x, "]"
    stop 1
end if
deallocate(x)
end program t
