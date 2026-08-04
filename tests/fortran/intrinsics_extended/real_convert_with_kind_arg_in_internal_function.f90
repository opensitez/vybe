! vybe-test: fortran/intrinsics_extended/real_convert_with_kind_arg_in_internal_function
! origin: languages/fortran/tests/fortran/test_intrinsics_extended.rs
program t
integer, parameter :: dp = kind(1.0d0)
if ((sample(7)) /= 7) then
    print *, "FAIL: want [7] got [", sample(7), "]"
    stop 1
end if
contains
pure function sample(s) result(r)
integer, intent(in) :: s
real(dp) :: r
r = real(s, dp)
end function sample
end program t
