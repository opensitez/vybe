! vybe-test: fortran/fortran2008/module_default_private
! origin: languages/fortran/tests/fortran/test_fortran2008.rs

module strict_mod
    implicit none
    private
    public :: visible

    integer :: hidden = 0
    integer, public :: visible = 42
end module strict_mod

program test
    use strict_mod
    print *, visible
end program test
