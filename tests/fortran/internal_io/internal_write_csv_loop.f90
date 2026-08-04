! vybe-test: fortran/internal_io/internal_write_csv_loop
! origin: languages/fortran/tests/fortran/test_internal_io.rs

program test
    character(len=40) :: row
    integer :: i
    do i = 1, 3
        write(row, '(I0, ",", I0, ",", I0)') i, i+1, i+2
        print *, trim(row)
    end do
end program test
