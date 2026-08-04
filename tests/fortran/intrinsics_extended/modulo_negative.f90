! vybe-test: fortran/intrinsics_extended/modulo_negative
! origin: languages/fortran/tests/fortran/test_intrinsics_extended.rs
program t
if ((modulo(-1, 5)) /= 4) then
    print *, "FAIL: want [4] got [", modulo(-1, 5), "]"
    stop 1
end if
end program t
