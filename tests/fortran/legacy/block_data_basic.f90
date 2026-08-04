! vybe-test: fortran/legacy/block_data_basic
! origin: languages/fortran/tests/fortran/test_legacy.rs

block data init_data
    integer :: x, y
    common /shared/ x, y
    data x /10/, y /20/
end block data init_data

program test
    integer :: x, y
    common /shared/ x, y
    print *, x + y
end program test
