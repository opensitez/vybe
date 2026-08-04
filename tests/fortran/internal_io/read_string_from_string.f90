! vybe-test: fortran/internal_io/read_string_from_string
! origin: languages/fortran/tests/fortran/test_internal_io.rs

program test
    character(len=20) :: buf = 'hello world'
    character(len=5) :: word
    read(buf, '(A5)') word
    print *, word
end program test
