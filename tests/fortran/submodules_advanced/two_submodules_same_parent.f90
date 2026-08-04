! vybe-test: fortran/submodules_advanced/two_submodules_same_parent
! origin: languages/fortran/tests/fortran/test_submodules_advanced.rs

module math_iface
    implicit none
    interface
        module function add(a, b) result(r)
            integer, intent(in) :: a, b
            integer :: r
        end function add
        module function sub(a, b) result(r)
            integer, intent(in) :: a, b
            integer :: r
        end function sub
    end interface
end module math_iface

submodule (math_iface) math_add
    implicit none
contains
    module function add(a, b) result(r)
        integer, intent(in) :: a, b
        integer :: r
        r = a + b
    end function add
end submodule math_add

submodule (math_iface) math_sub
    implicit none
contains
    module function sub(a, b) result(r)
        integer, intent(in) :: a, b
        integer :: r
        r = a - b
    end function sub
end submodule math_sub

program test
    use math_iface
    if ((add(10, 5)) /= 15) then
    print *, "FAIL: want [15] got [", add(10, 5), "]"
    stop 1
end if
    if ((sub(10, 5)) /= 5) then
    print *, "FAIL: want [5] got [", sub(10, 5), "]"
    stop 1
end if
end program test
