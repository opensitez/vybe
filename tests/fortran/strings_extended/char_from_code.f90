! vybe-test: fortran/strings_extended/char_from_code
! origin: languages/fortran/tests/fortran/test_strings_extended.rs
program t
character :: c
c = char(72)
print *, c
end program t
