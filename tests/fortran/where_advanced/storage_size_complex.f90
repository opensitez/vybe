! vybe-test: fortran/where_advanced/storage_size_complex
! origin: languages/fortran/tests/fortran/test_where_advanced.rs

program test
    complex :: c = (0., 0.)
    real :: r = 0.
    print *, storage_size(c) == 2 * storage_size(r)
end program test
