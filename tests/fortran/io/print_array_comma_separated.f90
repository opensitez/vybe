! vybe-test: fortran/io/print_array_comma_separated
! origin: languages/fortran/tests/fortran/test_io.rs

program test
    integer :: a(3) = [1,2,3]
    print '(3(I0, ","))', a
end program test
