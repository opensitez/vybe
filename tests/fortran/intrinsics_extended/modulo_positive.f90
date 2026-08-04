! vybe-test: fortran/intrinsics_extended/modulo_positive
! origin: languages/fortran/tests/fortran/test_intrinsics_extended.rs
program t
if ((modulo(10, 3)) /= 1) then
    print *, "FAIL: want [1] got [", modulo(10, 3), "]"
    stop 1
end if
end program t
