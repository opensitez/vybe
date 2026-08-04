! vybe-test: fortran/character/char_internal_io_29
! origin: languages/fortran/tests/fortran/test_character.rs
program p
character(len=20) :: buf
write(buf,'(A)') 'abc'
print *, trim(buf)
end program p
