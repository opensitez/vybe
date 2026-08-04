! vybe-test: fortran/associate_construct_extended/associate_shift_left
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
integer :: v = 3
associate (shifted => ishft(v, 1))
if ((shifted) /= 6) then
    print *, "FAIL: want [6] got [", shifted, "]"
    stop 1
end if
end associate
end program t
