! vybe-test: fortran/io_advanced/format_label_string
! origin: languages/fortran/tests/fortran/test_io_advanced.rs

program test
    character(len=5) :: s = 'hello'
300 format(A10)
    write(*, 300) s
end program test
