! vybe-test: fortran/intrinsics_extended/merge_basic
! origin: languages/fortran/tests/fortran/test_intrinsics_extended.rs
program t
if ((merge(1, 0, .true.)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, .true.), "]"
    stop 1
end if
end program t
