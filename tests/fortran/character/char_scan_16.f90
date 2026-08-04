! vybe-test: fortran/character/char_scan_16
! origin: languages/fortran/tests/fortran/test_character.rs
program p
print *, scan('abc123','0123456789')
end program p
