! vybe-test: fortran/where_advanced/where_three_elsewhere_clauses
! origin: languages/fortran/tests/fortran/test_where_advanced.rs

program test
    real :: t(6) = [-10., -1., 0., 1., 10., 100.]
    integer :: cat(6)
    where (t < -5.0)
        cat = 1
    elsewhere (t < 0.0)
        cat = 2
    elsewhere (t < 5.0)
        cat = 3
    elsewhere
        cat = 4
    end where
    print *, cat(1)
    print *, cat(2)
    print *, cat(4)
    print *, cat(6)
end program test
