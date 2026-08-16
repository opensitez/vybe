! vybe-test: fortran/array_transfer_between_kinds/array_transfer_between_kinds_real_to_int
! origin: languages/fortran/tests/fortran/test_array_transfer_between_kinds.rs

program array_transfer_between_kinds_real_to_int
    real :: source
    integer :: sink
    source = 2.0
    sink = transfer(source, sink)
    if ((sink) /= 1073741824) then
    print *, "FAIL: want [1073741824] got [", sink, "]"
    stop 1
end if
end program array_transfer_between_kinds_real_to_int
