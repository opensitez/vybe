! vybe-test: fortran/fortran2018_extended/out_of_range_two_hundred_exceeds_int8
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs
program t
integer :: x = 200
if ((out_of_range(x, 0_1)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", out_of_range(x, 0_1), "]"
    stop 1
end if
end program t
