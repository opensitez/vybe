! vybe-test: fortran/implicit_interfaces/implicit_interfaces_11
! origin: languages/fortran/tests/fortran/test_implicit_interfaces.rs
integer function f()

f = 9
end function f
program p
external f
integer :: f
integer :: x
x = f()
if (x /= 9) then
    print *, "FAIL: want [9] got [", x, "]"
    stop 1
end if
end program p
