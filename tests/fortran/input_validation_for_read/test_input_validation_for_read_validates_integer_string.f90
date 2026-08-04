! vybe-test: fortran/input_validation_for_read/test_input_validation_for_read_validates_integer_string
! origin: languages/fortran/tests/fortran/test_input_validation_for_read.rs

program test_input_validation_for_read
    character(len=8) :: src
    integer :: value
    integer :: status
    src = '42'
    read(src, *, iostat=status) value
    if ((status) /= 0) then
    print *, "FAIL: want [0] got [", status, "]"
    stop 1
end if
    if ((value) /= 42) then
    print *, "FAIL: want [42] got [", value, "]"
    stop 1
end if
end program test_input_validation_for_read
