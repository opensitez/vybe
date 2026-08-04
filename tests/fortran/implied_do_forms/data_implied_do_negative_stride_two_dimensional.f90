! vybe-test: fortran/implied_do_forms/data_implied_do_negative_stride_two_dimensional
! origin: languages/fortran/tests/fortran/test_implied_do_forms.rs

program t
    integer :: arr(3)
    data (arr(i), i = 5, 1, -2) /5, 3, 1/
    print *, arr(1)
end program t
