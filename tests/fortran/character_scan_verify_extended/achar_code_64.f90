! vybe-test: fortran/character_scan_verify_extended/achar_code_64
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=1) :: c
c = achar(64)
if ((ichar(c)) /= 64) then
    print *, "FAIL: want [64] got [", ichar(c), "]"
    stop 1
end if
end program t
