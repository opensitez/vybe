! vybe-test: fortran/control/ctrl_block_02
! origin: languages/fortran/tests/fortran/test_control.rs
program p
block
 integer :: x
 x = 7
 if ((x) /= 7) then
    print *, "FAIL: want [7] got [", x, "]"
    stop 1
end if
end block
end program p
