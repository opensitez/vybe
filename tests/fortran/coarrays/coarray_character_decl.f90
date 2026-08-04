! vybe-test: fortran/coarrays/coarray_character_decl
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    character(len=20) :: msg[*]
    msg = 'hello'
    print *, trim(msg)
end program test
