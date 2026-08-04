! vybe-test: fortran/format_io_extended/fmt_three_f_different_precision
! origin: languages/fortran/tests/fortran/test_format_io_extended.rs
program t
print '(F4.1, F5.2, F6.3)', 1.1, 2.22, 3.333
end program t
