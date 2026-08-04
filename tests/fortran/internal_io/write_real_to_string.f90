! vybe-test: fortran/internal_io/write_real_to_string
! origin: languages/fortran/tests/fortran/test_internal_io.rs

program test
    character(len=30) :: buf
    write(buf, *) 3.14159
    print *, trim(buf)
end program test
