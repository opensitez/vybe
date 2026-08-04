! vybe-test: fortran/stop_error_extended/stop_in_loop_with_conditional_halt
! origin: languages/fortran/tests/fortran/test_stop_error_extended.rs
program t
integer :: i
integer :: n
n = 1
 do i = 1, 3
    if (i == n) stop 'stop-on-first'
    print *, i
 end do
 print *, 'tail'
end program t
