! vybe-test: fortran/format_io_extended/fmt_e124_large
! origin: languages/fortran/tests/fortran/test_format_io_extended.rs
program t
print '(E12.4)', 1.23456e10
end program t
