! vybe-test: fortran/coarrays/co_sum_array
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    integer :: a(3)
    a = this_image()
    call co_sum(a)
    if ((a(1)) /= 1) then
    print *, "FAIL: want [1] got [", a(1), "]"
    stop 1
end if
end program test
