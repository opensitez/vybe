! vybe-test: fortran/submodules_advanced/three_submodules_same_parent_runtime
! origin: languages/fortran/tests/fortran/test_submodules_advanced.rs

module ops_iface
    implicit none
    interface
        module function mul(a, b) result(r)
            integer, intent(in) :: a, b
            integer :: r
        end function mul
        module function div(a, b) result(r)
            integer, intent(in) :: a, b
            integer :: r
        end function div
        module function modulo_op(a, b) result(r)
            integer, intent(in) :: a, b
            integer :: r
        end function modulo_op
    end interface
end module ops_iface

submodule (ops_iface) ops_mul
contains
    module function mul(a, b) result(r)
        integer, intent(in) :: a, b
        integer :: r
        r = a * b
    end function mul
end submodule ops_mul

submodule (ops_iface) ops_div
contains
    module function div(a, b) result(r)
        integer, intent(in) :: a, b
        integer :: r
        r = a / b
    end function div
end submodule ops_div

submodule (ops_iface) ops_mod
contains
    module function modulo_op(a, b) result(r)
        integer, intent(in) :: a, b
        integer :: r
        r = mod(a, b)
    end function modulo_op
end submodule ops_mod

program test
    use ops_iface
    if ((mul(6, 7)) /= 42) then
    print *, "FAIL: want [42] got [", mul(6, 7), "]"
    stop 1
end if
    if ((div(20, 4)) /= 5) then
    print *, "FAIL: want [5] got [", div(20, 4), "]"
    stop 1
end if
    if ((modulo_op(17, 5)) /= 2) then
    print *, "FAIL: want [2] got [", modulo_op(17, 5), "]"
    stop 1
end if
end program test
