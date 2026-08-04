! vybe-test: fortran/intrinsics_extended/real_convert_with_kind_arg
! origin: languages/fortran/tests/fortran/test_intrinsics_extended.rs
program t
integer, parameter :: dp = kind(1.0d0)
if ((real(7, dp)) /= 7) then
    print *, "FAIL: want [7] got [", real(7, dp), "]"
    stop 1
end if
end program t
