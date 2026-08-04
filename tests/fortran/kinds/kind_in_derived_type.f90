! vybe-test: fortran/kinds/kind_in_derived_type
! origin: languages/fortran/tests/fortran/test_kinds.rs

program test
    integer, parameter :: wp = 8
    type :: HighPrec
        real(kind=wp) :: value
    end type HighPrec
    type(HighPrec) :: h
    h%value = 1.23456789012345_wp
    print *, h%value
end program test
