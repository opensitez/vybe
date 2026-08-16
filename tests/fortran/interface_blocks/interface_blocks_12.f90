! vybe-test: fortran/interface_blocks/interface_blocks_12
! origin: languages/fortran/tests/fortran/test_interface_blocks.rs
integer function f(a, b, scale)
integer, intent(in) :: a, b
integer, intent(in), optional :: scale
f = a + b
if (present(scale)) f = f * scale
end function f
program t
interface
integer function f(a, b, scale)
integer, intent(in) :: a, b
integer, intent(in), optional :: scale
end function f
end interface
if (f(2, 3) /= 5) then
    print *, "FAIL: want [5] got [", f(2, 3), "]"
    stop 1
end if
if (f(2, 3, 4) /= 20) then
    print *, "FAIL: want [20] got [", f(2, 3, 4), "]"
    stop 1
end if
end program t
