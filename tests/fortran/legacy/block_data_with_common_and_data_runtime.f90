! vybe-test: fortran/legacy/block_data_with_common_and_data_runtime
! origin: languages/fortran/tests/fortran/test_legacy.rs

block data init_data
    integer :: x, y
    common /shared/ x, y
    data x /10/, y /20/
end block data init_data

program test
    integer :: x, y
    common /shared/ x, y
    if ((x + y) /= 30) then
    print *, "FAIL: want [30] got [", x + y, "]"
    stop 1
end if
end program test
