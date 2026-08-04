! vybe-test: fortran/character/char_format_write_28
! origin: languages/fortran/tests/fortran/test_character.rs
program p
character(len=5) :: s='abc'
write(*,'(A)') s
end program p
