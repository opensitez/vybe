! vybe-test: fortran/complex/cmplx_runtime_real_and_imag_parts
! origin: languages/fortran/tests/fortran/test_complex.rs
program t
  complex :: z
  z = cmplx(3.0, 4.0)
  if ((real(z)) /= 3) then
    print *, "FAIL: want [3] got [", real(z), "]"
    stop 1
end if
  if ((aimag(z)) /= 4) then
    print *, "FAIL: want [4] got [", aimag(z), "]"
    stop 1
end if
end program t
