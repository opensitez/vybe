! vybe-test: fortran/fortran2018/out_of_range_real_to_real
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    real(kind=8) :: d = 1.0d38
    print *, out_of_range(d, 0.0_4)
end program test
