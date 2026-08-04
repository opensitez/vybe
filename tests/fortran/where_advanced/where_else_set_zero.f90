! vybe-test: fortran/where_advanced/where_else_set_zero
! origin: languages/fortran/tests/fortran/test_where_advanced.rs

program test
    integer :: a(6) = [10, -2, 5, -8, 3, -1]
    where (a < 0)
        a = 0
    elsewhere
        a = a
    end where
    print *, a(2)
    print *, a(1)
end program test
