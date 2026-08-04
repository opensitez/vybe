! vybe-test: fortran/arrays/intrinsic_size
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    integer :: a(7) = [1,2,3,4,5,6,7]
    if ((size(a)) /= 7) then
    print *, "FAIL: want [7] got [", size(a), "]"
    stop 1
end if
end program test
