! vybe-test: fortran/input_validation_for_read/test_input_validation_for_read_parses_real_value
! origin: languages/fortran/tests/fortran/test_input_validation_for_read.rs

program test_input_validation_for_read
    character(len=16) :: src
    real :: value
    integer :: status
    src = '1.25e1'
    read(src, *, iostat=status) value
    if ((status) /= 0) then
    print *, "FAIL: want [0] got [", status, "]"
    stop 1
end if
    if ((nint(value)) /= 13) then
    print *, "FAIL: want [13] got [", nint(value), "]"
    stop 1
end if
end program test_input_validation_for_read
