! vybe-test: fortran/arrays/where_basic
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    integer :: a(5) = [1, -2, 3, -4, 5]
    integer :: b(5)
    where (a > 0)
        b = a
    elsewhere
        b = 0
    end where
    print *, b(1)
end program test
