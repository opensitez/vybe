! vybe-test: fortran/io_file_position/fio_read_iostat_eof_on_second_scratch_read
! origin: languages/fortran/tests/fortran/test_io_file_position.rs

program t
    integer :: n, ios
    open(10, status='scratch')
    write(10, '(I0)') 42
    rewind(10)
    read(10, *, iostat=ios) n
    read(10, *, iostat=ios) n
    if (ios /= 0) print *, 1
    close(10)
end program t
