! vybe-test: fortran/interface_blocks/interface_blocks_14
! origin: languages/fortran/tests/fortran/test_interface_blocks.rs
logical function has_value(v)
integer, intent(in) :: v
has_value = v /= 0
end function has_value
program t
interface
logical function has_value(v)
integer, intent(in) :: v
end function has_value
end interface
if (.not. has_value(4)) then
    print *, "FAIL: want [true] got [", has_value(4), "]"
    stop 1
end if
if (has_value(0)) then
    print *, "FAIL: want [false] got [", has_value(0), "]"
    stop 1
end if
end program t
