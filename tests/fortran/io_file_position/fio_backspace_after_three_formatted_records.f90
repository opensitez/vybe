! vybe-test: fortran/io_file_position/fio_backspace_after_three_formatted_records
! origin: languages/fortran/tests/fortran/test_io_file_position.rs

program t
    open(10, status='scratch')
    write(10, '(I0)') 1
    write(10, '(I0)') 2
    write(10, '(I0)') 3
    backspace(10)
    close(10)
end program t
