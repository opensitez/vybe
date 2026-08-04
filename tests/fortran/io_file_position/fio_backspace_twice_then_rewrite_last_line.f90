! vybe-test: fortran/io_file_position/fio_backspace_twice_then_rewrite_last_line
! origin: languages/fortran/tests/fortran/test_io_file_position.rs

program t
    open(10, status='scratch')
    write(10, '(A)') 'alpha'
    write(10, '(A)') 'beta'
    write(10, '(A)') 'gamma'
    backspace(10)
    backspace(10)
    write(10, '(A)') 'delta'
    close(10)
end program t
