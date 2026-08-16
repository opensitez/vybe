! vybe-test: fortran/interface_blocks/interface_blocks_11
! origin: languages/fortran/tests/fortran/test_interface_blocks.rs
subroutine s(x, y)
integer, intent(inout) :: x
integer, intent(in) :: y
x = x + y
end subroutine s
program t
interface
subroutine s(x, y)
integer, intent(inout) :: x
integer, intent(in) :: y
end subroutine s
end interface
integer :: v
v = 5
call s(v, 7)
if (v /= 12) then
    print *, "FAIL: want [12] got [", v, "]"
    stop 1
end if
end program t
