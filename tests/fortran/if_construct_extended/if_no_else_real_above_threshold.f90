! vybe-test: fortran/if_construct_extended/if_no_else_real_above_threshold
! origin: languages/fortran/tests/fortran/test_if_construct_extended.rs
program t
real :: r = 3.5
if (r > 3.0) then
if (trim("above") /= "above") then
    print *, "FAIL: want [above] got [", "above", "]"
    stop 1
end if
end if
end program t
