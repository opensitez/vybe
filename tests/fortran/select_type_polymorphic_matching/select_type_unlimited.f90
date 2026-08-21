! vybe-test: fortran/select_type_polymorphic_matching/select_type_unlimited
! origin: languages/fortran/tests/fortran/test_fortran2003.rs

program test
    class(*), allocatable :: val
    allocate(integer :: val)
    select type(val)
    type is (integer)
        print *, 'integer'
    type is (real)
        print *, 'real'
    class default
        print *, 'other'
    end select
end program test
