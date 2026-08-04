! vybe-test: fortran/legacy/data_character
! origin: languages/fortran/tests/fortran/test_legacy.rs

program test
    character(len=5) :: s
    data s /'hello'/
    print *, s
end program test
