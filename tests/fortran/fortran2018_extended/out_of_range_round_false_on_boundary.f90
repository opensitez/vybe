! vybe-test: fortran/fortran2018_extended/out_of_range_round_false_on_boundary
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs

program t
    real :: x = 127.6
    print *, out_of_range(x, 0_1, round=.false.)
end program t
