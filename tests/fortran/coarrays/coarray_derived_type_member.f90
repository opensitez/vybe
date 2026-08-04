! vybe-test: fortran/coarrays/coarray_derived_type_member
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    type :: Shared
        integer :: value
    end type Shared
    type(Shared) :: obj[*]
    obj%value = this_image()
    sync all
    print *, obj%value
end program test
