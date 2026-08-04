! vybe-test: fortran/stop_error_extended/error_stop_code_guard_not_taken
! origin: languages/fortran/tests/fortran/test_stop_error_extended.rs
program t
logical :: err = .false.
if (trim('check') /= "check") then
    print *, "FAIL: want [check] got [", 'check', "]"
    stop 1
end if
if (err) error stop 7
if (trim('ok') /= "ok") then
    print *, "FAIL: want [ok] got [", 'ok', "]"
    stop 1
end if
end program t
