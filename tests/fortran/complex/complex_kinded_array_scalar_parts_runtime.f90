! vybe-test: fortran/complex/complex_kinded_array_scalar_parts_runtime
! origin: languages/fortran/tests/fortran/test_complex.rs
program t
  integer, parameter :: dp = kind(1.0d0)
  complex(dp) :: x(2)
  x(1) = cmplx(2.0_dp, 4.0_dp, dp)
  x(2) = cmplx(6.0_dp, 8.0_dp, dp)
  if ((nint(real(x(1)))) /= 2) then
    print *, "FAIL: want [2] got [", nint(real(x(1))), "]"
    stop 1
end if
  if ((nint(aimag(x(1)))) /= 4) then
    print *, "FAIL: want [4] got [", nint(aimag(x(1))), "]"
    stop 1
end if
  if ((nint(real(x(2)))) /= 6) then
    print *, "FAIL: want [6] got [", nint(real(x(2))), "]"
    stop 1
end if
  if ((nint(aimag(x(2)))) /= 8) then
    print *, "FAIL: want [8] got [", nint(aimag(x(2))), "]"
    stop 1
end if
end program t
