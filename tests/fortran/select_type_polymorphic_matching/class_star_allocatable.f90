! vybe-test: fortran/select_type_polymorphic_matching/class_star_allocatable
! origin: languages/fortran/tests/fortran/test_fortran2003.rs

program test
    class(*), allocatable :: obj
    allocate(integer :: obj)
    print *, "ok"
end program test
