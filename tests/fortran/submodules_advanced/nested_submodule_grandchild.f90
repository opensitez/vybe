! vybe-test: fortran/submodules_advanced/nested_submodule_grandchild
! origin: languages/fortran/tests/fortran/test_submodules_advanced.rs

module base_mod
    implicit none
    interface
        module function compute(x) result(r)
            integer, intent(in) :: x
            integer :: r
        end function compute
    end interface
end module base_mod

submodule (base_mod) child_mod
    implicit none
    interface
        module function helper(x) result(r)
            integer, intent(in) :: x
            integer :: r
        end function helper
    end interface
end submodule child_mod

submodule (base_mod:child_mod) grandchild_mod
    implicit none
contains
    module function compute(x) result(r)
        integer, intent(in) :: x
        integer :: r
        r = helper(x) * 2
    end function compute

    module function helper(x) result(r)
        integer, intent(in) :: x
        integer :: r
        r = x + 1
    end function helper
end submodule grandchild_mod

program test
    use base_mod
    if ((compute(5)) /= 12) then
    print *, "FAIL: want [12] got [", compute(5), "]"
    stop 1
end if
end program test
