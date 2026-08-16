! vybe-test: fortran/interface_blocks/interface_blocks_15
! origin: languages/fortran/tests/fortran/test_interface_blocks.rs
character(len=4) function aschar(i)
integer, intent(in) :: i
write(aschar, '(i4)') i
end function aschar
program t
interface
character(len=4) function aschar(i)
integer, intent(in) :: i
end function aschar
end interface
if (adjustl(aschar(42)) /= "42  ") then
    print *, "FAIL: want [42] got [", adjustl(aschar(42)), "]"
    stop 1
end if
end program t
