! vybe-test: fortran/complex/complex_array_scalar_division_runtime
! origin: languages/fortran/tests/fortran/test_complex.rs
program t
  complex :: x(2)
  x(1) = cmplx(2.0, 4.0)
  x(2) = cmplx(6.0, 8.0)
  x = x / 2.0
  if ((real(x(1))) /= 1) then
    print *, "FAIL: want [1] got [", real(x(1)), "]"
    stop 1
end if
  if ((aimag(x(1))) /= 2) then
    print *, "FAIL: want [2] got [", aimag(x(1)), "]"
    stop 1
end if
  if ((real(x(2))) /= 3) then
    print *, "FAIL: want [3] got [", real(x(2)), "]"
    stop 1
end if
  if ((aimag(x(2))) /= 4) then
    print *, "FAIL: want [4] got [", aimag(x(2)), "]"
    stop 1
end if
end program t
