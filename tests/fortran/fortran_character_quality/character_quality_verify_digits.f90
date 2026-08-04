! vybe-test: fortran/fortran_character_quality/character_quality_verify_digits
! origin: languages/fortran/tests/fortran/test_fortran_character_quality.rs

program character_quality_verify_digits
    character(len=20) :: source
    source = 'abc123def'
    if ((verify(source, '0123456789')) /= 1) then
    print *, "FAIL: want [1] got [", verify(source, '0123456789'), "]"
    stop 1
end if
    if ((verify(source, 'abcdef')) /= 4) then
    print *, "FAIL: want [4] got [", verify(source, 'abcdef'), "]"
    stop 1
end if
end program character_quality_verify_digits
