! vybe-test: fortran/initialization/init_alloc_char_13
! origin: languages/fortran/tests/fortran/test_initialization.rs
program p
character(len=:),allocatable::s
allocate(character(len=3)::s)
s='abc'
end program p
