! vybe-test: fortran/where_advanced/storage_size_runtime_basics
! origin: languages/fortran/tests/fortran/test_where_advanced.rs

program test
    integer :: i = 0
    real :: r = 0.0
    real(kind=8) :: d = 0.0d0
    if ((storage_size(i)) /= 32) then
    print *, "FAIL: want [32] got [", storage_size(i), "]"
    stop 1
end if
    if ((storage_size(r)) /= 32) then
    print *, "FAIL: want [32] got [", storage_size(r), "]"
    stop 1
end if
    if ((storage_size(d)) /= 64) then
    print *, "FAIL: want [64] got [", storage_size(d), "]"
    stop 1
end if
end program test
