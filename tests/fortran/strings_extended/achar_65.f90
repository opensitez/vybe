! vybe-test: fortran/strings_extended/achar_65
! origin: languages/fortran/tests/fortran/test_strings_extended.rs
program t
character :: c
c = achar(65)
print *, c
end program t
