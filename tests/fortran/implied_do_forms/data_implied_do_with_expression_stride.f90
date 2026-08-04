! vybe-test: fortran/implied_do_forms/data_implied_do_with_expression_stride
! origin: languages/fortran/tests/fortran/test_implied_do_forms.rs

program t
    integer :: v(4)
    data (v(i), i = 2, 8, 2 + 0) /2, 4, 6, 8/
    print *, v(4)
end program t
