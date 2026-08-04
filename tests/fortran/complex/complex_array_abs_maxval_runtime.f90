! vybe-test: fortran/complex/complex_array_abs_maxval_runtime
! origin: languages/fortran/tests/fortran/test_complex.rs
program t
  complex :: x(2)
  x(1) = cmplx(3.0, 4.0)
  x(2) = cmplx(1.0, 2.0)
  if ((maxval(abs(x))) /= 5) then
    print *, "FAIL: want [5] got [", maxval(abs(x)), "]"
    stop 1
end if
end program t
