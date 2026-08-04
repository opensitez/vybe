! vybe-test: fortran/fortran2018/fortran_2018_reduce_runtime_sum_and_mask
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program t
    integer :: a(6) = [1, 2, 3, 4, 5, 6]
    integer :: total
    logical :: mask(6) = [.true., .false., .true., .false., .true., .false.]
    total = reduce(a, operator(+))
    if ((total) /= 21) then
    print *, "FAIL: want [21] got [", total, "]"
    stop 1
end if
    if ((reduce(a, operator(+), mask=mask)) /= 9) then
    print *, "FAIL: want [9] got [", reduce(a, operator(+), mask=mask), "]"
    stop 1
end if
end program t
