! vybe-test: fortran/complex_extended/real_of_sum_runtime
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: a, b, c
a = cmplx(2.0, 3.0)
b = cmplx(4.0, 5.0)
c = a + b
if ((nint(real(c))) /= 6) then
    print *, "FAIL: want [6] got [", nint(real(c)), "]"
    stop 1
end if
end program t
