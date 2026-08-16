! vybe-test: fortran/interface_blocks/interface_blocks_13
! origin: languages/fortran/tests/fortran/test_interface_blocks.rs
subroutine copy(src, dst)
integer, intent(in) :: src(:)
integer, intent(out) :: dst(:)
dst = src * 3
end subroutine copy
program t
interface
subroutine copy(src, dst)
integer, intent(in) :: src(:)
integer, intent(out) :: dst(:)
end subroutine copy
end interface
integer :: a(3), b(3)
a = [1, 2, 3]
b = 0
call copy(a, b)
if (sum(b) /= 18) then
    print *, "FAIL: want [18] got [", sum(b), "]"
    stop 1
end if
end program t
