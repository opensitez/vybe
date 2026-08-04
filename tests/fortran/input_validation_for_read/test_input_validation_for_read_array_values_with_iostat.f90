! vybe-test: fortran/input_validation_for_read/test_input_validation_for_read_array_values_with_iostat
! origin: languages/fortran/tests/fortran/test_input_validation_for_read.rs

program test_input_validation_for_read
    character(len=32) :: src
    integer :: values(3)
    integer :: status
    src = '1 2 3'
    read(src, *, iostat=status) values
    if ((status) /= 0) then
    print *, "FAIL: want [0] got [", status, "]"
    stop 1
end if
    if ((sum(values)) /= 6) then
    print *, "FAIL: want [6] got [", sum(values), "]"
    stop 1
end if
end program test_input_validation_for_read
