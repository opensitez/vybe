! vybe-test: fortran/pointer_alloc_extended/alloc_logical_scalar_value
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
logical, allocatable :: flag
allocate(flag)
flag = .true.
if (flag .neqv. .true.) then
    print *, "FAIL: want [true] got [", flag, "]"
    stop 1
end if
deallocate(flag)
end program t
