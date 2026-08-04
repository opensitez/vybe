! vybe-test: fortran/internal_io/write_multiple_values
! origin: languages/fortran/tests/fortran/test_internal_io.rs

program test
    character(len=30) :: buf
    write(buf, *) 1, 2, 3
    print *, trim(buf)
end program test
