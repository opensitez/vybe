! vybe-test: fortran/format_io_extended/fmt_f_variable_real
! origin: languages/fortran/tests/fortran/test_format_io_extended.rs
program t
real :: x = 1.25
print '(F6.2)', x
end program t
