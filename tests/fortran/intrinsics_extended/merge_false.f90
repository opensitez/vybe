! vybe-test: fortran/intrinsics_extended/merge_false
! origin: languages/fortran/tests/fortran/test_intrinsics_extended.rs
program t
if ((merge(1, 0, .false.)) /= 0) then
    print *, "FAIL: want [0] got [", merge(1, 0, .false.), "]"
    stop 1
end if
end program t
