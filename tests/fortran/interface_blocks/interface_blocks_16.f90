! vybe-test: fortran/interface_blocks/interface_blocks_16
! origin: languages/fortran/tests/fortran/test_interface_blocks.rs
integer function custom_add(a, b)
integer, intent(in) :: a, b
custom_add = a * 10 + b
end function custom_add
program t
interface operator(.custom.)
integer function custom_add(a, b)
integer, intent(in) :: a, b
end function custom_add
end interface
integer :: p, q, r
p = 3
q = 4
r = p .custom. q
if (r /= 34) then
    print *, "FAIL: want [34] got [", r, "]"
    stop 1
end if
end program t
