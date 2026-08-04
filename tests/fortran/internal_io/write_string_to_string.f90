! vybe-test: fortran/internal_io/write_string_to_string
! origin: languages/fortran/tests/fortran/test_internal_io.rs

program test
    character(len=20) :: buf
    write(buf, '(A)') 'hello'
    print *, trim(buf)
end program test
