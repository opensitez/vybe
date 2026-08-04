! vybe-test: fortran/format_io_extended/fmt_label_write_multi
! origin: languages/fortran/tests/fortran/test_format_io_extended.rs
program t
integer :: n = 3
real :: r = 1.5
write(*, 500) n, r
500 format(I0, F5.1)
end program t
