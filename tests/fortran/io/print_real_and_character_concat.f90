! vybe-test: fortran/io/print_real_and_character_concat
! origin: languages/fortran/tests/fortran/test_io.rs

program test
    character(len=8) :: buf
    real :: x
    x = 2.5
    write(buf, '(A,F4.1)') "x=", x
    print *, trim(buf)
end program test
