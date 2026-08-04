! vybe-test: fortran/format_io_extended/fmt_a_character_variable
! origin: languages/fortran/tests/fortran/test_format_io_extended.rs
program t
character(len=5) :: s = 'world'
print '(A)', s
end program t
