! vybe-test: fortran/io/write_character_padding_retained
! origin: languages/fortran/tests/fortran/test_io.rs

program test
    write(*, '(A4)') 'abc'
end program test
