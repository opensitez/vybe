! vybe-test: fortran/io_file_position/fio_scratch_close_status_delete
! origin: languages/fortran/tests/fortran/test_io_file_position.rs

program t
    open(10, status='scratch')
    write(10, '(I0)') 1
    close(10, status='delete')
end program t
