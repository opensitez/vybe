! vybe-test: fortran/statement_f77_legacy_compat/statement_f77_legacy_compat_format_write_variants
! origin: languages/fortran/tests/fortran/test_statement_f77_legacy_compat.rs
program statement_f77_legacy_compat_format_write_variants
integer i
i = 4
print 10, i
10          format (I1)
end program statement_f77_legacy_compat_format_write_variants
