! vybe-test: fortran/associate_construct_extended/associate_array_last_element
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
integer :: a(5) = [11, 22, 33, 44, 55]
associate (tail => a(5))
if ((tail) /= 55) then
    print *, "FAIL: want [55] got [", tail, "]"
    stop 1
end if
end associate
end program t
