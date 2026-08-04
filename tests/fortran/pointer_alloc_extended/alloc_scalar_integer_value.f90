! vybe-test: fortran/pointer_alloc_extended/alloc_scalar_integer_value
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
integer, allocatable :: n
allocate(n)
n = 99
if ((n) /= 99) then
    print *, "FAIL: want [99] got [", n, "]"
    stop 1
end if
deallocate(n)
end program t
