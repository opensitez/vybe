! vybe-test: fortran/pointer_alloc_extended/alloc_real_three_first_element
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
real, allocatable :: r(:)
allocate(r(3))
r = [1.5, 2.5, 3.5]
if (abs((r(1)) - 1.5) > 1.0e-6) then
    print *, "FAIL: want [1.5] got [", r(1), "]"
    stop 1
end if
deallocate(r)
end program t
