! vybe-test: fortran/io/formatted_array_print_ignores_unused_repeat_descriptors
! origin: languages/fortran/tests/fortran/test_io.rs

program test
    real(8) :: a(2)
    a(1) = 1.25d0
    a(2) = 2.5d0
    print '(4f10.4)', a
end program test
