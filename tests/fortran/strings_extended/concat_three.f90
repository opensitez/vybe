! vybe-test: fortran/strings_extended/concat_three
! origin: languages/fortran/tests/fortran/test_strings_extended.rs

program test
    character(len=5) :: a = 'Hello'
    character(len=2) :: b = ', '
    character(len=5) :: c = 'World'
    character(len=15) :: s
    s = trim(a) // trim(b) // trim(c)
    print *, trim(s)
end program test
