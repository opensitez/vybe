! vybe-test: fortran/fortran2018_extended/out_of_range_zero_fits_int16
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs
program t
integer :: x = 0
if ((out_of_range(x, 0_2)) .neqv. .false.) then
    print *, "FAIL: want [false] got [", out_of_range(x, 0_2), "]"
    stop 1
end if
end program t
