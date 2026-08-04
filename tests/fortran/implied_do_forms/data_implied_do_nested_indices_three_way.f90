! vybe-test: fortran/implied_do_forms/data_implied_do_nested_indices_three_way
! origin: languages/fortran/tests/fortran/test_implied_do_forms.rs

program t
    integer :: vol(2,2,2)
    data (((vol(i, j, k), k = 1, 2), j = 1, 2), i = 1, 2) /1,2,3,4,5,6,7,8/
    print *, vol(2, 2, 2)
end program t
