! vybe-test: fortran/internal_io/write_int_to_string
! origin: languages/fortran/tests/fortran/test_internal_io.rs

program test
    character(len=20) :: buf
    write(buf, *) 42
    print *, trim(buf)
end program test
