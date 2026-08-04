! vybe-test: fortran/pointer_alloc_extended/pointer_array_third_element
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
integer, target :: data(5) = [3, 6, 9, 12, 15]
integer, pointer :: slice(:)
slice => data
if ((slice(3)) /= 9) then
    print *, "FAIL: want [9] got [", slice(3), "]"
    stop 1
end if
end program t
