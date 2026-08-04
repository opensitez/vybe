! vybe-test: fortran/fortran2018/out_of_range_integer_overflow
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    integer(kind=8) :: big = 1000000000000_8
    print *, out_of_range(big, 0_2)
end program test
