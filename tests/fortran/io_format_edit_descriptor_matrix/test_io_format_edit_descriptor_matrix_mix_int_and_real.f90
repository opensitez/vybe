! vybe-test: fortran/io_format_edit_descriptor_matrix/test_io_format_edit_descriptor_matrix_mix_int_and_real
! origin: languages/fortran/tests/fortran/test_io_format_edit_descriptor_matrix.rs

program test_io_format_edit_descriptor_matrix
    integer :: n
    real :: x
    n = 12
    x = 4.5
    print '(I4,1X,F4.1)', n, x
end program test_io_format_edit_descriptor_matrix
