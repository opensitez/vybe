! vybe-test: fortran/internal_io/write_formatted_real
! origin: languages/fortran/tests/fortran/test_internal_io.rs

program test
    character(len=15) :: buf
    write(buf, '(F8.3)') 3.14159
    print *, trim(buf)
end program test
