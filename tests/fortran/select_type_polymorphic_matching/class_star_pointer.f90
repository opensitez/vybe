! vybe-test: fortran/select_type_polymorphic_matching/class_star_pointer
! origin: languages/fortran/tests/fortran/test_fortran2003.rs

program test
    class(*), pointer :: p => null()
    integer, target :: x = 42
    p => x
    print *, "ok"
end program test
