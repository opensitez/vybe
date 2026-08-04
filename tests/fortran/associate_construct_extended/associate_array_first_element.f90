! vybe-test: fortran/associate_construct_extended/associate_array_first_element
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
integer :: a(5) = [11, 22, 33, 44, 55]
associate (head => a(1))
if ((head) /= 11) then
    print *, "FAIL: want [11] got [", head, "]"
    stop 1
end if
end associate
end program t
