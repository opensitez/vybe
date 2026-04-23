use super::helpers::compile_ok;

// ── Multiple submodules for the same parent ───────────────────

#[test] fn two_submodules_same_parent() {
    compile_ok(r#"
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
    print *, add(10, 5)
    print *, sub(10, 5)
end program test
"#);
}

#[test] fn three_submodules_same_parent() {
    compile_ok(r#"
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
    print *, mul(6, 7)
    print *, div(20, 4)
    print *, modulo_op(17, 5)
end program test
"#);
}

// ── Nested submodule (grandchild) ─────────────────────────────

#[test] fn nested_submodule_grandchild() {
    compile_ok(r#"
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
    print *, compute(5)
end program test
"#);
}

// ── Submodule using parent module types ───────────────────────

#[test] fn submodule_uses_parent_type() {
    compile_ok(r#"
module geometry_iface
    implicit none
    type :: Point
        real :: x, y
    end type Point
    interface
        module function distance(a, b) result(d)
            type(Point), intent(in) :: a, b
            real :: d
        end function distance
    end interface
end module geometry_iface

submodule (geometry_iface) geometry_impl
    implicit none
contains
    module function distance(a, b) result(d)
        type(Point), intent(in) :: a, b
        real :: d
        d = sqrt((a%x - b%x)**2 + (a%y - b%y)**2)
    end function distance
end submodule geometry_impl

program test
    use geometry_iface
    type(Point) :: p1, p2
    p1 = Point(0.0, 0.0)
    p2 = Point(3.0, 4.0)
    print *, distance(p1, p2)
end program test
"#);
}

// ── Submodule with private module variables ───────────────────

#[test] fn submodule_private_state() {
    compile_ok(r#"
module counter_iface
    implicit none
    interface
        module subroutine increment()
        end subroutine increment
        module function get_count() result(n)
            integer :: n
        end function get_count
    end interface
end module counter_iface

submodule (counter_iface) counter_impl
    implicit none
    integer :: count = 0
contains
    module subroutine increment()
        count = count + 1
    end subroutine increment

    module function get_count() result(n)
        integer :: n
        n = count
    end function get_count
end submodule counter_impl

program test
    use counter_iface
    call increment()
    call increment()
    call increment()
    print *, get_count()
end program test
"#);
}

// ── Submodule with multiple procedures + helpers ──────────────

#[test] fn submodule_with_internal_helpers() {
    compile_ok(r#"
module stats_iface
    implicit none
    interface
        module function mean(a) result(m)
            real, intent(in) :: a(:)
            real :: m
        end function mean
        module function variance(a) result(v)
            real, intent(in) :: a(:)
            real :: v
        end function variance
    end interface
end module stats_iface

submodule (stats_iface) stats_impl
    implicit none
contains
    module function mean(a) result(m)
        real, intent(in) :: a(:)
        real :: m
        m = sum(a) / real(size(a))
    end function mean

    module function variance(a) result(v)
        real, intent(in) :: a(:)
        real :: v
        real :: m
        m = mean(a)
        v = sum((a - m)**2) / real(size(a))
    end function variance
end submodule stats_impl

program test
    use stats_iface
    real :: data(5) = [1.0, 2.0, 3.0, 4.0, 5.0]
    print *, mean(data)
    print *, variance(data)
end program test
"#);
}

// ── Submodule with subroutine and function ────────────────────

#[test] fn submodule_mixed_sub_and_func() {
    compile_ok(r#"
module io_iface
    implicit none
    interface
        module subroutine print_vec(v)
            real, intent(in) :: v(:)
        end subroutine print_vec
        module function dot(u, v) result(d)
            real, intent(in) :: u(:), v(:)
            real :: d
        end function dot
    end interface
end module io_iface

submodule (io_iface) io_impl
    implicit none
contains
    module subroutine print_vec(v)
        real, intent(in) :: v(:)
        integer :: i
        do i = 1, size(v)
            print *, v(i)
        end do
    end subroutine print_vec

    module function dot(u, v) result(d)
        real, intent(in) :: u(:), v(:)
        real :: d
        d = sum(u * v)
    end function dot
end submodule io_impl

program test
    use io_iface
    real :: u(3) = [1.0, 2.0, 3.0]
    real :: v(3) = [4.0, 5.0, 6.0]
    print *, dot(u, v)
end program test
"#);
}

// ── Submodule with DEFAULT ACCESSIBILITY ─────────────────────

#[test] fn submodule_with_protected_parent_var() {
    compile_ok(r#"
module cfg_iface
    implicit none
    integer, protected :: max_size = 100
    interface
        module subroutine set_max(n)
            integer, intent(in) :: n
        end subroutine set_max
    end interface
end module cfg_iface

submodule (cfg_iface) cfg_impl
    implicit none
contains
    module subroutine set_max(n)
        integer, intent(in) :: n
        max_size = n
    end subroutine set_max
end submodule cfg_impl

program test
    use cfg_iface
    print *, max_size
    call set_max(200)
    print *, max_size
end program test
"#);
}

// ── Submodule USE of parent exported generic ──────────────────

#[test] fn submodule_with_generic_interface() {
    compile_ok(r#"
module generic_iface
    implicit none
    interface norm
        module function norm_real(a) result(r)
            real, intent(in) :: a(:)
            real :: r
        end function norm_real
        module function norm_dbl(a) result(r)
            real(kind=8), intent(in) :: a(:)
            real(kind=8) :: r
        end function norm_dbl
    end interface norm
end module generic_iface

submodule (generic_iface) generic_impl
    implicit none
contains
    module function norm_real(a) result(r)
        real, intent(in) :: a(:)
        real :: r
        r = sqrt(sum(a**2))
    end function norm_real

    module function norm_dbl(a) result(r)
        real(kind=8), intent(in) :: a(:)
        real(kind=8) :: r
        r = sqrt(sum(a**2))
    end function norm_dbl
end submodule generic_impl

program test
    use generic_iface
    real :: v(3) = [3.0, 4.0, 0.0]
    print *, norm(v)
end program test
"#);
}
