! vybe-test: fortran/do_loops/bare_do_loop_exits
! origin: languages/fortran/tests/fortran/test_do_loops.rs
program t
integer :: i
i = 0
do
i = i + 1
if (i >= 3) exit
end do
if ((i) /= 3) then
    print *, "FAIL: want [3] got [", i, "]"
    stop 1
end if
end program t
