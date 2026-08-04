! vybe-test: fortran/character_scan_verify_extended/repeat_once_is_identity
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
if (trim(repeat('ok', 1)) /= "ok") then
    print *, "FAIL: want [ok] got [", repeat('ok', 1), "]"
    stop 1
end if
end program t
