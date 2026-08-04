! vybe-test: fortran/arrays/where_set_to_zero
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    integer :: v(6) = [1, -2, 3, -4, 5, -6]
    where (v < 0)
        v = 0
    end where
    print *, v(1)
end program test
