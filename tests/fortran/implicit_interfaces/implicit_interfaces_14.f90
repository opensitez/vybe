! vybe-test: fortran/implicit_interfaces/implicit_interfaces_14
! origin: languages/fortran/tests/fortran/test_implicit_interfaces.rs
complex function f(a, b)
real :: a, b
f = cmplx(a, b)
end function f
program p
external f
complex :: f
complex :: z
z = f(1.0, 2.0)
if (nint(real(z) + aimag(z)) /= 3) then
    print *, "FAIL: want [3] got [", nint(real(z) + aimag(z)), "]"
    stop 1
end if
end program p
