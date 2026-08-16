! vybe-test: fortran/explicit_interfaces/explicit_interfaces_04
! origin: languages/fortran/tests/fortran/test_explicit_interfaces.rs
integer function f()
f = 3
end function f
program t
interface
integer function f()
end function f
end interface
if (f() /= 3) then
    print *, "FAIL: want [3] got [", f(), "]"
    stop 1
end if
end program t
