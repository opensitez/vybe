! vybe-test: fortran/associate_construct_extended/associate_block_then_outer_print
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
integer :: val = 9
associate (v => val)
if ((v) /= 9) then
    print *, "FAIL: want [9] got [", v, "]"
    stop 1
end if
end associate
if ((val) /= 9) then
    print *, "FAIL: want [9] got [", val, "]"
    stop 1
end if
end program t
