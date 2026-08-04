! vybe-test: fortran/complex_extended/aimag_of_product_nint
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: a, b, p
a = cmplx(2.0, 0.0)
b = cmplx(0.0, 3.0)
p = a * b
if ((nint(aimag(p))) /= 6) then
    print *, "FAIL: want [6] got [", nint(aimag(p)), "]"
    stop 1
end if
end program t
