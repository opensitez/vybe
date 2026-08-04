! vybe-test: fortran/internal_io/format_hex_as_string
! origin: languages/fortran/tests/fortran/test_internal_io.rs

program test
    character(len=10) :: s
    write(s, '(Z8)') 255
    print *, trim(s)
end program test
