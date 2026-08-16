! vybe-test: fortran/math_f2008/bessel_yn_order_0
! origin: languages/fortran/tests/fortran/test_math_f2008.rs
program t
  if (abs((bessel_yn(0, 1.0)) - (0.0882569626)) > 1.000000e-06) then
      print *, "FAIL: want [0.0882569626] got [", bessel_yn(0, 1.0), "]"
      stop 1
  end if
end program t
