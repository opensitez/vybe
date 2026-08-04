! vybe-test: fortran/array_transfer_between_kinds/array_transfer_between_kinds_logical_to_int
! origin: languages/fortran/tests/fortran/test_array_transfer_between_kinds.rs

program array_transfer_between_kinds_logical_to_int
    logical :: ok
    integer :: mark
    ok = .true.
    mark = transfer(ok, mark)
    if ((mark) /= 1) then
    print *, "FAIL: want [1] got [", mark, "]"
    stop 1
end if
end program array_transfer_between_kinds_logical_to_int
