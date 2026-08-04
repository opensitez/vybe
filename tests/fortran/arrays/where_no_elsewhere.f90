! vybe-test: fortran/arrays/where_no_elsewhere
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    real :: a(4) = [1.0, -2.0, 3.0, -4.0]
    where (a < 0.0)
        a = 0.0
    end where
    print *, a(1)
end program test
