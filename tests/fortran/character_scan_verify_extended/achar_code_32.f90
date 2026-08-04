! vybe-test: fortran/character_scan_verify_extended/achar_code_32
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=1) :: c
c = achar(32)
if ((ichar(c)) /= 32) then
    print *, "FAIL: want [32] got [", ichar(c), "]"
    stop 1
end if
end program t
