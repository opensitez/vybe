! vybe-test: fortran/format_io_extended/fmt_label_write_integer
! origin: languages/fortran/tests/fortran/test_format_io_extended.rs
program t
integer :: i = 7
write(*, 100) i
100 format(I5)
end program t
