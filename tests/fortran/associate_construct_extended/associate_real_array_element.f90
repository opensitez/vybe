! vybe-test: fortran/associate_construct_extended/associate_real_array_element
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
real :: r(3) = [1.5, 2.5, 3.5]
associate (mid => r(2))
if ((int(mid * 2.0)) /= 5) then
    print *, "FAIL: want [5] got [", int(mid * 2.0), "]"
    stop 1
end if
end associate
end program t
