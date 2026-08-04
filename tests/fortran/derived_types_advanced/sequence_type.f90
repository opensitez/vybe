! vybe-test: fortran/derived_types_advanced/sequence_type
! origin: languages/fortran/tests/fortran/test_derived_types_advanced.rs

program test
    type :: Packed
        sequence
        integer :: a
        real :: b
        logical :: c
    end type Packed
    type(Packed) :: p
    p%a = 1
    p%b = 2.0
    p%c = .true.
    print *, p%a
end program test
