! vybe-test: fortran/implicit_interfaces/implicit_interfaces_02
! origin: languages/fortran/tests/fortran/test_implicit_interfaces.rs
integer function f()
f = 5
end function f
program p
integer f
external f
if (f() /= 5) then
    print *, "FAIL: want [5] got [", f(), "]"
    stop 1
end if
end program p
