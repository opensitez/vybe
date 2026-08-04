! vybe-test: fortran/fortran2018_extended/out_of_range_integer_to_smaller_kind_negative
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs

program t
    integer :: x = -200
    print *, out_of_range(x, 0_1)
end program t
