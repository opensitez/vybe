! vybe-test: fortran/array_transfer_between_kinds/array_transfer_between_kinds_int_to_real
! origin: languages/fortran/tests/fortran/test_array_transfer_between_kinds.rs

program array_transfer_between_kinds_int_to_real
    integer :: source
    real :: sink
    source = 1
    sink = transfer(source, sink)
    if ((int(sink)) /= 0) then
    print *, "FAIL: want [0] got [", int(sink), "]"
    stop 1
end if
end program array_transfer_between_kinds_int_to_real
