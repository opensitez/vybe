! vybe-test: fortran/internal_io/build_csv_line
! origin: languages/fortran/tests/fortran/test_internal_io.rs

program test
    character(len=60) :: line
    integer :: a = 1, b = 2, c = 3
    write(line, '(I0, A, I0, A, I0)') a, ',', b, ',', c
    print *, trim(line)
end program test
