! vybe-test: fortran/internal_io/iostat_on_internal_read
! origin: languages/fortran/tests/fortran/test_internal_io.rs

program test
    character(len=5) :: buf = 'abc'
    integer :: n, ios
    read(buf, *, iostat=ios) n
    if (ios /= 0) then
        print *, 'parse error'
    end if
end program test
