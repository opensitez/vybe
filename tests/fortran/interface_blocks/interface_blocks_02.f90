! vybe-test: fortran/interface_blocks/interface_blocks_02
! origin: languages/fortran/tests/fortran/test_interface_blocks.rs
subroutine s(x)
integer::x
x = x + 1
end subroutine s
program t

interface
subroutine s(x)
integer::x
end subroutine s
end interface
integer :: v
v = 4
call s(v)
if (v /= 5) then
    print *, "FAIL: want [5] got [", v, "]"
    stop 1
end if
end program t
