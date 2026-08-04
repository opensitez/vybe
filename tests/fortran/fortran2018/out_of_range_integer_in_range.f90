! vybe-test: fortran/fortran2018/out_of_range_integer_in_range
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    integer(kind=4) :: x = 100
    print *, out_of_range(x, 0_2)
end program test
