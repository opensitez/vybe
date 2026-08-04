! vybe-test: fortran/io_file_position/fio_inquire_file_exist_after_replace_create
! origin: languages/fortran/tests/fortran/test_io_file_position.rs

program t
    logical :: exists
    open(10, file='fio_inquire_exist.dat', status='replace')
    write(10, '(I0)') 1
    close(10)
    inquire(file='fio_inquire_exist.dat', exist=exists)
    print *, exists
end program t
