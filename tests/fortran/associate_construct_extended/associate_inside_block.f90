! vybe-test: fortran/associate_construct_extended/associate_inside_block
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
integer :: x = 8
block
integer :: y
y = x + 2
associate (z => y * 3)
if ((z) /= 30) then
    print *, "FAIL: want [30] got [", z, "]"
    stop 1
end if
end associate
end block
end program t
