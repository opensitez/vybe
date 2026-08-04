! vybe-test: fortran/format_io_extended/fmt_record_advance_with_slash
! origin: languages/fortran/tests/fortran/test_format_io_extended.rs
program t
print '(A, /, A)', 'top', 'bottom'
end program t
