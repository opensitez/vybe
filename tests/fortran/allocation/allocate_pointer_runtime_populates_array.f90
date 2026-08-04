! vybe-test: fortran/allocation/allocate_pointer_runtime_populates_array
! origin: languages/fortran/tests/fortran/test_allocation.rs
program t
integer, pointer :: p(:)
allocate(p(3))
p = [1, 2, 3]
if ((p(1)) /= 1) then
    print *, "FAIL: want [1] got [", p(1), "]"
    stop 1
end if
if ((p(3)) /= 3) then
    print *, "FAIL: want [3] got [", p(3), "]"
    stop 1
end if
deallocate(p)
end program t
