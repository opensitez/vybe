! vybe-test: fortran/io_advanced/read_iostat
! origin: languages/fortran/tests/fortran/test_io_advanced.rs

program test
    integer :: n, ios
    read(*, *, iostat=ios) n
    if (ios /= 0) then
        print *, 'read error'
    else
        print *, n
    end if
end program test
