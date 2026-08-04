! vybe-test: fortran/internal_io/write_double_to_string
! origin: languages/fortran/tests/fortran/test_internal_io.rs

program test
    character(len=30) :: buf
    real(kind=8) :: d = 2.718281828d0
    write(buf, '(D20.12)') d
    print *, trim(buf)
end program test
