! vybe-test: fortran/io_advanced/format_label
! origin: languages/fortran/tests/fortran/test_io_advanced.rs

program test
    integer :: i = 7
    write(*, 100) i
100 format(I5)
end program test
