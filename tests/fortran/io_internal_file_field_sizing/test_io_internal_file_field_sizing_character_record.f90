! vybe-test: fortran/io_internal_file_field_sizing/test_io_internal_file_field_sizing_character_record
! origin: languages/fortran/tests/fortran/test_io_internal_file_field_sizing.rs

program test_io_internal_file_field_sizing
    integer :: n
    character(len=12) :: name
    name = 'fortran_test'
    inquire(iolength=n) name
    if ((n) /= 12) then
    print *, "FAIL: want [12] got [", n, "]"
    stop 1
end if
end program test_io_internal_file_field_sizing
