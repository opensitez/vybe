! vybe-test: fortran/explicit_interfaces/explicit_interfaces_03
! origin: languages/fortran/tests/fortran/test_explicit_interfaces.rs
real function f(x)
real::x
f = x * 2.0
end function f
program t
interface
real function f(x)
real::x
end function f
end interface
if (nint(f(1.5)) /= 3) then
    print *, "FAIL: want [3] got [", nint(f(1.5)), "]"
    stop 1
end if
end program t
