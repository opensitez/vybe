! vybe-test: fortran/character_arrays/char_array_param_init
! origin: languages/fortran/tests/fortran/test_character_arrays.rs

program test
    character(len=3), parameter :: days(7) = &
        ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
    print *, days(1)
    print *, days(5)
end program test
