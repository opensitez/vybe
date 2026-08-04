! vybe-test: fortran/io_file_position/fio_inquire_unit_access_and_form_sequential
! origin: languages/fortran/tests/fortran/test_io_file_position.rs

program t
    character(len=16) :: access, form
    open(10, file='fio_inquire_attrs.dat', status='replace')
    inquire(unit=10, access=access, form=form)
    close(10, status='delete')
    print *, trim(access)
    print *, trim(form)
end program t
