! vybe-test: fortran/legacy_data_extended/block_data_two_common_blocks_runtime
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs

block data setup
    integer :: ix
    real :: rx
    common /ints/ ix
    common /reals/ rx
    data ix /42/, rx /3.5/
end block data setup

program t
    integer :: ix
    real :: rx
    common /ints/ ix
    common /reals/ rx
    if ((ix) /= 42) then
    print *, "FAIL: want [42] got [", ix, "]"
    stop 1
end if
    if (abs((rx) - 3.5) > 1.0e-6) then
    print *, "FAIL: want [3.5] got [", rx, "]"
    stop 1
end if
end program t
