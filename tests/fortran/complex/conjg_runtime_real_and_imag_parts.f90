! vybe-test: fortran/complex/conjg_runtime_real_and_imag_parts
! origin: languages/fortran/tests/fortran/test_complex.rs
program t
  complex :: z, c
  z = cmplx(3.0, 4.0)
  c = conjg(z)
  if ((real(c)) /= 3) then
    print *, "FAIL: want [3] got [", real(c), "]"
    stop 1
end if
  if ((aimag(c)) /= -4) then
    print *, "FAIL: want [-4] got [", aimag(c), "]"
    stop 1
end if
end program t
