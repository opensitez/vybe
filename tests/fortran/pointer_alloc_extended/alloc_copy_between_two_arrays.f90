! vybe-test: fortran/pointer_alloc_extended/alloc_copy_between_two_arrays
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
integer, allocatable :: src(:), dst(:)
src = [7, 8, 9]
dst = src
if ((dst(1)) /= 7) then
    print *, "FAIL: want [7] got [", dst(1), "]"
    stop 1
end if
if ((dst(3)) /= 9) then
    print *, "FAIL: want [9] got [", dst(3), "]"
    stop 1
end if
end program t
