! vybe-test: fortran/associate_construct_extended/associate_sequential_two_blocks
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
integer :: x = 3, y = 7
associate (a => x)
if ((a) /= 3) then
    print *, "FAIL: want [3] got [", a, "]"
    stop 1
end if
end associate
associate (b => y)
if ((b) /= 7) then
    print *, "FAIL: want [7] got [", b, "]"
    stop 1
end if
end associate
end program t
