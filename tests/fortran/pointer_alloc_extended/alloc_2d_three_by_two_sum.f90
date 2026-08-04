! vybe-test: fortran/pointer_alloc_extended/alloc_2d_three_by_two_sum
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
integer, allocatable :: m(:,:)
allocate(m(3, 2))
m = 1
if ((sum(m)) /= 6) then
    print *, "FAIL: want [6] got [", sum(m), "]"
    stop 1
end if
if ((size(m)) /= 6) then
    print *, "FAIL: want [6] got [", size(m), "]"
    stop 1
end if
deallocate(m)
end program t
