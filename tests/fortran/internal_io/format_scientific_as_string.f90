! vybe-test: fortran/internal_io/format_scientific_as_string
! origin: languages/fortran/tests/fortran/test_internal_io.rs

program test
    character(len=20) :: s
    write(s, '(E12.4)') 1.23456e10
    print *, trim(s)
end program test
