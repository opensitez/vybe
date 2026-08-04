! vybe-test: fortran/intrinsics/sum_array_runtime
! origin: languages/fortran/tests/fortran/test_intrinsics.rs

program test
    integer :: a(4) = [1, 2, 3, 4]
    if ((sum(a)) /= 10) then
    print *, "FAIL: want [10] got [", sum(a), "]"
    stop 1
end if
end program test
