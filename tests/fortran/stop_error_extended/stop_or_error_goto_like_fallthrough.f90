! vybe-test: fortran/stop_error_extended/stop_or_error_goto_like_fallthrough
! origin: languages/fortran/tests/fortran/test_stop_error_extended.rs
program t
integer :: code
code = 0
if (code /= 0) stop code
if (code == 0) then
    if (trim('ok') /= "ok") then
    print *, "FAIL: want [ok] got [", 'ok', "]"
    stop 1
end if
else
    error stop 'bad'
end if
if (trim('done') /= "done") then
    print *, "FAIL: want [done] got [", 'done', "]"
    stop 1
end if
end program t
