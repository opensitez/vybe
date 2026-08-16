! vybe-test: fortran/implicit_interfaces/implicit_interfaces_06
! origin: languages/fortran/tests/fortran/test_implicit_interfaces.rs
real function f()
f = 2.5
end function f
program p
external f
real :: f
real :: x
x = f()
if (nint(x * 2) /= 5) then
    print *, "FAIL: want [5] got [", nint(x * 2), "]"
    stop 1
end if
end program p
