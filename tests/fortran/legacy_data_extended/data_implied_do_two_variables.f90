! vybe-test: fortran/legacy_data_extended/data_implied_do_two_variables
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs

program t
    integer :: x, y, z
    data x /1/, y /2/, z /3/
    print *, x + y + z
end program t
