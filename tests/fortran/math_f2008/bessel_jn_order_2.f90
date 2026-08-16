! vybe-test: fortran/math_f2008/bessel_jn_order_2
! origin: languages/fortran/tests/fortran/test_math_f2008.rs
program t
  if (abs((bessel_jn(2, 1.0)) - (0.114903487)) > 1.149035e-06) then
      print *, "FAIL: want [0.114903487] got [", bessel_jn(2, 1.0), "]"
      stop 1
  end if
end program t
