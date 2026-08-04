! vybe-test: fortran/internal_io/write_formatted_int
! origin: languages/fortran/tests/fortran/test_internal_io.rs

program test
    character(len=10) :: buf
    write(buf, '(I5)') 42
    print *, trim(buf)
end program test
