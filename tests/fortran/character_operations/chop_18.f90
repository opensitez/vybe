! vybe-test: fortran/character_operations/chop_18
! origin: languages/fortran/tests/fortran/test_character_operations.rs
program p
character(len=20) :: buf
write(buf,'(A)') 'abc'
print *, trim(buf)
end program p
