! vybe-test: fortran/intrinsics_extended/sign_neg_to_pos
! origin: languages/fortran/tests/fortran/test_intrinsics_extended.rs
program t
if ((sign(-5, 1)) /= 5) then
    print *, "FAIL: want [5] got [", sign(-5, 1), "]"
    stop 1
end if
end program t
