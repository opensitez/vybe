! vybe-test: fortran/format_io_extended/fmt_multi_i_and_f
! origin: languages/fortran/tests/fortran/test_format_io_extended.rs
program t
integer :: n = 7
real :: x = 2.5
print '(I0, F6.2)', n, x
end program t
