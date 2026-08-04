! vybe-test: fortran/associate_construct_extended/associate_scalar_real_rename
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
real :: x = 3.5
associate (r => x)
if ((int(r * 2.0)) /= 7) then
    print *, "FAIL: want [7] got [", int(r * 2.0), "]"
    stop 1
end if
end associate
end program t
