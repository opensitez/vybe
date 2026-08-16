! vybe-test: fortran/interface_blocks/interface_blocks_19
! origin: languages/fortran/tests/fortran/test_interface_blocks.rs
subroutine f(x)
integer, pointer :: x
x = 12
end subroutine f
program t
interface
subroutine f(x)
integer, pointer :: x
end subroutine f
end interface
integer, pointer :: p
integer, target :: v
v = 0
p => v
call f(p)
if (v /= 12) then
    print *, "FAIL: want [12] got [", v, "]"
    stop 1
end if
end program t
