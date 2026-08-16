! vybe-test: fortran/implicit_interfaces/implicit_interfaces_12
! origin: languages/fortran/tests/fortran/test_implicit_interfaces.rs
real function f(a, n)
real :: a
integer :: n
f = a + real(n)
end function f
program p
external f
real :: f
real :: x, y
y = 1.5
x = f(y, 2)
if (nint(x * 2) /= 7) then
    print *, "FAIL: want [7] got [", nint(x * 2), "]"
    stop 1
end if
end program p
