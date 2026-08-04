! vybe-test: fortran/derived_types_advanced/class_unlimited_polymorphic
! origin: languages/fortran/tests/fortran/test_derived_types_advanced.rs

program test
    class(*), pointer :: p => null()
    print *, "ok"
end program test
