! vybe-test: fortran/associate_construct_extended/associate_multi_three_scalars
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
integer :: a = 1, b = 2, c = 3
associate (x => a, y => b, z => c)
if ((x + y + z) /= 6) then
    print *, "FAIL: want [6] got [", x + y + z, "]"
    stop 1
end if
end associate
end program t
