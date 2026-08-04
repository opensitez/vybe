! vybe-test: fortran/io_file_position/fio_inquire_unit_opened_after_scratch_open
! origin: languages/fortran/tests/fortran/test_io_file_position.rs

program t
    logical :: opened
    open(10, status='scratch')
    inquire(unit=10, opened=opened)
    close(10)
    print *, opened
end program t
