! vybe-test: fortran/legacy_data_extended/common_two_named_blocks
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs
program t
integer :: ix
real :: rx
common /ints/ ix
common /reals/ rx
ix = 7
rx = 2.5
if ((ix) /= 7) then
    print *, "FAIL: want [7] got [", ix, "]"
    stop 1
end if
if (abs((rx) - 2.5) > 1.0e-6) then
    print *, "FAIL: want [2.5] got [", rx, "]"
    stop 1
end if
end program t
