! vybe-test: fortran/select_type_polymorphic_matching/select_type_polymorphic_matching_no_match_no_default_no_output
! origin: languages/fortran/tests/fortran/test_select_type_polymorphic_matching.rs

program select_type_polymorphic_matching_no_match_no_default_no_output
    class(*), allocatable :: value
    allocate(real :: value)
    select type (value)
    type is (integer)
        print *, 1
    class is (character(*))
        print *, 2
    end select
end program select_type_polymorphic_matching_no_match_no_default_no_output
