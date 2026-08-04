! vybe-test: fortran/implied_do_forms/data_implied_do_with_parameterized_bounds
! origin: languages/fortran/tests/fortran/test_implied_do_forms.rs

program t
    integer, parameter :: n = 3
    integer :: out(n)
    data (out(i), i = 1, n) /10,20,30/
    print *, out(2)
end program t
