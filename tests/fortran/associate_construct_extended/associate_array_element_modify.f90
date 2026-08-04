! vybe-test: fortran/associate_construct_extended/associate_array_element_modify
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
integer :: a(4) = [1, 2, 3, 4]
associate (slot => a(2))
slot = 99
end associate
if ((a(2)) /= 99) then
    print *, "FAIL: want [99] got [", a(2), "]"
    stop 1
end if
end program t
