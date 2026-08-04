! vybe-test: fortran/format_io_extended/fmt_5x_spacing
! origin: languages/fortran/tests/fortran/test_format_io_extended.rs
program t
print '(A, 5X, A)', 'L', 'R'
end program t
