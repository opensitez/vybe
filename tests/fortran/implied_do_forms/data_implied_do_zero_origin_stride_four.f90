! vybe-test: fortran/implied_do_forms/data_implied_do_zero_origin_stride_four
! origin: languages/fortran/tests/fortran/test_implied_do_forms.rs

program t
    integer :: slots(3)
    data (slots(i), i = 0, 8, 4) /0, 4, 8/
    print *, slots(2)
end program t
