! vybe-test: fortran/input_validation_for_read/test_input_validation_for_read_parses_logical_from_string
! origin: languages/fortran/tests/fortran/test_input_validation_for_read.rs

program test_input_validation_for_read
    character(len=8) :: src
    logical :: value
    integer :: status
    src = '.true.'
    read(src, *, iostat=status) value
    if ((status) /= 0) then
    print *, "FAIL: want [0] got [", status, "]"
    stop 1
end if
    if (trim(value) /= ".TRUE.") then
    print *, "FAIL: want [.TRUE.] got [", value, "]"
    stop 1
end if
end program test_input_validation_for_read
