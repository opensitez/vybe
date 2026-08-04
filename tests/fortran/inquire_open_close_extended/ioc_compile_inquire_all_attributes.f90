! vybe-test: fortran/inquire_open_close_extended/ioc_compile_inquire_all_attributes
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs

program t
    logical :: opened, named, exist
    integer :: num, rec, ios
    character(len=20) :: acc, frm, nm
    open(10, file='ioc_ext_inq.dat', status='replace')
    inquire(unit=10, opened=opened, number=num, named=named, name=nm, &
            access=acc, form=frm, rec=rec, iostat=ios)
    close(10, status='delete')
    print *, 1
end program t
