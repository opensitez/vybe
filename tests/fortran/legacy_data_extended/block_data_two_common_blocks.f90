! vybe-test: fortran/legacy_data_extended/block_data_two_common_blocks
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
    print *, ix
    print *, rx
end program t
