! vybe-test: fortran/coarrays/coarray_scalar_decl
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    integer :: x[*]
    x = 42
    if ((x) /= 42) then
    print *, "FAIL: want [42] got [", x, "]"
    stop 1
end if
end program test
