! vybe-test: fortran/math_f2008/bessel_yn_order_2
! origin: languages/fortran/tests/fortran/test_math_f2008.rs
program t
  if (abs((bessel_yn(2, 1.0)) - (-1.65068257)) > 1.650683e-05) then
      print *, "FAIL: want [-1.65068257] got [", bessel_yn(2, 1.0), "]"
      stop 1
  end if
end program t
