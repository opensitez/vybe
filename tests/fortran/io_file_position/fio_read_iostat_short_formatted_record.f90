! vybe-test: fortran/io_file_position/fio_read_iostat_short_formatted_record
! origin: languages/fortran/tests/fortran/test_io_file_position.rs

program t
    integer :: a, b, ios
    open(10, status='scratch')
    write(10, '(I0)') 7
    rewind(10)
    read(10, '(2I4)', iostat=ios) a, b
    if (ios /= 0) print *, ios
    close(10)
end program t
