! vybe-test: fortran/out_of_range/out_of_range_with_round
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    real :: x = 127.6
    print *, out_of_range(x, 0_1, round=.true.)
end program test
