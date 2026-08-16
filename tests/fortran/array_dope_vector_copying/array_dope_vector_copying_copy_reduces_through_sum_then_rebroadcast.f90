! vybe-test: fortran/array_dope_vector_copying/array_dope_vector_copying_copy_reduces_through_sum_then_rebroadcast
! origin: languages/fortran/tests/fortran/test_array_dope_vector_copying.rs

program t
    integer, allocatable :: source(:)
    integer, allocatable :: target(:)
    integer :: total
    source = (/ 1, 2, 3, 4, 5, 6 /)
    total = sum(source)
    target = (/ total /)
    if ((size(target)) /= 1) then
    print *, "FAIL: want [1] got [", size(target), "]"
    stop 1
end if
    if ((sum(target)) /= 21) then
    print *, "FAIL: want [21] got [", sum(target), "]"
    stop 1
end if
    if ((target(1)) /= 21) then
    print *, "FAIL: want [21] got [", target(1), "]"
    stop 1
end if
end program t
