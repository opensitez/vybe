! vybe-test: fortran/format_io_extended/fmt_e_negative_exponent
! origin: languages/fortran/tests/fortran/test_format_io_extended.rs
program t
print '(E11.3)', -4.5e3
end program t
