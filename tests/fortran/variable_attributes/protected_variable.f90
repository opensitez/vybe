! vybe-test: fortran/variable_attributes/protected_variable
! origin: languages/fortran/tests/fortran/test_fortran2003.rs

module prot_mod
    implicit none
    integer, protected :: counter = 0
contains
    subroutine increment()
        counter = counter + 1
    end subroutine increment
end module prot_mod

program test
    use prot_mod
    call increment()
    print *, counter
end program test
