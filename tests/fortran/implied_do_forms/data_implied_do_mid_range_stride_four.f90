! vybe-test: fortran/implied_do_forms/data_implied_do_mid_range_stride_four
! origin: languages/fortran/tests/fortran/test_implied_do_forms.rs

program t
    integer :: buf(3)
    data (buf(i), i = 3, 11, 4) /11, 15, 19/
    print *, buf(3)
end program t
