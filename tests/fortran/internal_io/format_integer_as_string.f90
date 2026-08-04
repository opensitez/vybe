! vybe-test: fortran/internal_io/format_integer_as_string
! origin: languages/fortran/tests/fortran/test_internal_io.rs

program test
    character(len=10) :: s
    integer :: n = 255
    write(s, '(I0)') n
    print *, trim(s)
end program test
