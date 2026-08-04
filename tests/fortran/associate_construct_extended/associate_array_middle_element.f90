! vybe-test: fortran/associate_construct_extended/associate_array_middle_element
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
integer :: a(5) = [11, 22, 33, 44, 55]
associate (mid => a(3))
if ((mid) /= 33) then
    print *, "FAIL: want [33] got [", mid, "]"
    stop 1
end if
end associate
end program t
