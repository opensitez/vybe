! vybe-test: fortran/complex_extended/array_slice_second_real_part
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: a(4)
a(1) = cmplx(10.0, 0.0)
a(2) = cmplx(20.0, 0.0)
a(3) = cmplx(30.0, 0.0)
a(4) = cmplx(40.0, 0.0)
if ((nint(real(a(2:4)(2)))) /= 30) then
    print *, "FAIL: want [30] got [", nint(real(a(2:4)(2))), "]"
    stop 1
end if
end program t
