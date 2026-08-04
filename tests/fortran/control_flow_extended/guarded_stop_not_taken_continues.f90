! vybe-test: fortran/control_flow_extended/guarded_stop_not_taken_continues
! origin: languages/fortran/tests/fortran/test_control_flow_extended.rs
program t
logical :: ok = .true.
if (.not. ok) stop 1
if (trim('ok') /= "ok") then
    print *, "FAIL: want [ok] got [", 'ok', "]"
    stop 1
end if
end program t
