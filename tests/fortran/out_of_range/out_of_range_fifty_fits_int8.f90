! vybe-test: fortran/out_of_range/out_of_range_fifty_fits_int8
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs
program t
integer :: x = 50
if ((out_of_range(x, 0_1)) .neqv. .false.) then
    print *, "FAIL: want [false] got [", out_of_range(x, 0_1), "]"
    stop 1
end if
end program t
