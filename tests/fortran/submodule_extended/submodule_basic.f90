! vybe-test: fortran/submodule_extended/submodule_basic
! origin: languages/fortran/tests/fortran/test_fortran2008.rs

module parent_mod
    implicit none
    interface
        module function compute(x) result(r)
            integer, intent(in) :: x
            integer :: r
        end function compute
    end interface
end module parent_mod

submodule (parent_mod) parent_mod_impl
    implicit none
contains
    module function compute(x) result(r)
        integer, intent(in) :: x
        integer :: r
        r = x * x
    end function compute
end submodule parent_mod_impl

program test
    use parent_mod
    print *, compute(5)
end program test
