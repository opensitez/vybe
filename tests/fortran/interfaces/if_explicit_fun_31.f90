! vybe-test: fortran/interfaces/if_explicit_fun_31
! origin: languages/fortran/tests/fortran/test_interfaces.rs
real function f(x)
real :: x
f = x * 2.0
end function f
program t
interface
real function f(x)
real :: x
end function f
end interface
if (abs(f(1.5) - 3.0) > 1.0e-6) then
    print *, "FAIL: want [3.0] got [", f(1.5), "]"
    stop 1
end if
end program t
