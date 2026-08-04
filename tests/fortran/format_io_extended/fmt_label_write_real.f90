! vybe-test: fortran/format_io_extended/fmt_label_write_real
! origin: languages/fortran/tests/fortran/test_format_io_extended.rs
program t
real :: x = 2.718
write(*, 200) x
200 format(F8.3)
end program t
