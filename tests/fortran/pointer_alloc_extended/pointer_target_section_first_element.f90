! vybe-test: fortran/pointer_alloc_extended/pointer_target_section_first_element
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
integer, target :: series(6) = [10, 20, 30, 40, 50, 60]
integer, pointer :: window(:)
window => series(2:4)
if ((window(1)) /= 20) then
    print *, "FAIL: want [20] got [", window(1), "]"
    stop 1
end if
if ((size(window)) /= 3) then
    print *, "FAIL: want [3] got [", size(window), "]"
    stop 1
end if
end program t
