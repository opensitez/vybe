! vybe-test: fortran/associate_construct_extended/associate_array_2d_element
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
integer :: m(2,3) = reshape([1,2,3,4,5,6], [2,3])
associate (cell => m(2,1))
if ((cell) /= 2) then
    print *, "FAIL: want [2] got [", cell, "]"
    stop 1
end if
end associate
end program t
