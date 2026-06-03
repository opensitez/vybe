use super::helpers::{compile_ok, run_prints};

// ── Module with constants and functions ───────────────────────

#[test]
fn module_private_public() {
    compile_ok(
        r#"
module mymod
    implicit none
    private
    public :: get_value
    integer :: secret = 42
contains
    function get_value() result(v)
        integer :: v
        v = secret
    end function get_value
end module mymod

program test
    use mymod
    print *, get_value()
end program test
"#,
    );
}

#[test]
fn module_public_vars() {
    compile_ok(
        r#"
module config
    implicit none
    integer, public :: max_size = 100
    real, public :: tolerance = 1.0e-6
end module config

program test
    use config
    print *, max_size
end program test
"#,
    );
}

#[test]
fn module_use_rename() {
    compile_ok(
        r#"
module mathlib
    implicit none
contains
    function square(x) result(r)
        integer, intent(in) :: x
        integer :: r
        r = x * x
    end function square
end module mathlib

program test
    use mathlib, sq => square
    print *, sq(5)
end program test
"#,
    );
}

#[test]
fn module_multiple_use() {
    compile_ok(
        r#"
module constants
    real, parameter :: PI = 3.14159
end module constants

module helpers
    implicit none
contains
    function double(x) result(r)
        real, intent(in) :: x
        real :: r
        r = x * 2.0
    end function double
end module helpers

program test
    use constants
    use helpers
    print *, double(PI)
end program test
"#,
    );
}

// ── Module with derived types ─────────────────────────────────

#[test]
fn module_exports_type() {
    let out = run_prints(
        r#"
module geometry
    implicit none
    type :: Vector2D
        real :: x, y
    end type Vector2D
contains
    function length(v) result(r)
        type(Vector2D), intent(in) :: v
        real :: r
        r = sqrt(v%x**2 + v%y**2)
    end function length
end module geometry

program test
    use geometry
    type(Vector2D) :: v
    v%x = 3.0
    v%y = 4.0
    print *, length(v)
end program test
"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn module_type_with_procedure() {
    let out = run_prints(
        r#"
module animals
    implicit none
    type :: Dog
        character(len=20) :: name
    contains
        procedure :: speak
    end type Dog
contains
    subroutine speak(self)
        class(Dog), intent(in) :: self
        print *, 'Woof! I am ' // trim(self%name)
    end subroutine speak
end module animals

program test
    use animals
    type(Dog) :: d
    d%name = 'Rex'
    call d%speak()
end program test
"#,
    );
    assert_eq!(out, vec!["Woof! I am Rex"]);
}

// ── Module with SAVE attribute ────────────────────────────────

#[test]
fn module_save_variable() {
    compile_ok(
        r#"
module counter_mod
    implicit none
    integer, save :: count = 0
contains
    subroutine increment()
        count = count + 1
    end subroutine increment
    function get_count() result(c)
        integer :: c
        c = count
    end function get_count
end module counter_mod

program test
    use counter_mod
    call increment()
    call increment()
    call increment()
    print *, get_count()
end program test
"#,
    );
}

// ── USE, ONLY with multiple items ─────────────────────────────

#[test]
fn use_only_multiple() {
    compile_ok(
        r#"
module stuff
    integer :: a = 1, b = 2, c = 3
end module stuff

program test
    use stuff, only: a, c
    print *, a + c
end program test
"#,
    );
}

// ── INTERFACE blocks ──────────────────────────────────────────

#[test]
fn interface_explicit() {
    compile_ok(
        r#"
program test
    interface
        function square(x) result(r)
            integer, intent(in) :: x
            integer :: r
        end function square
    end interface
    print *, "ok"
end program test
"#,
    );
}

#[test]
fn interface_operator() {
    let out = run_prints(
        r#"
module vectors
    implicit none
    type :: Vec
        real :: x, y
    end type Vec
    interface operator(+)
        module procedure add_vecs
    end interface
contains
    function add_vecs(a, b) result(c)
        type(Vec), intent(in) :: a, b
        type(Vec) :: c
        c%x = a%x + b%x
        c%y = a%y + b%y
    end function add_vecs
end module vectors

program test
    use vectors
    type(Vec) :: v1, v2, v3
    v1 = Vec(1.0, 2.0)
    v2 = Vec(3.0, 4.0)
    v3 = v1 + v2
    print *, v3%x
end program test
"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn interface_assignment() {
    compile_ok(
        r#"
module conv
    implicit none
    interface assignment(=)
        module procedure int_to_real
    end interface
contains
    subroutine int_to_real(r, i)
        real, intent(out) :: r
        integer, intent(in) :: i
        r = real(i)
    end subroutine int_to_real
end module conv

program test
    print *, "ok"
end program test
"#,
    );
}

// ── Operator overloading ─────────────────────────────────────

#[test]
fn operator_plus_type() {
    let out = run_prints(
        r#"
module complex_mod
    implicit none
    type :: MyComplex
        real :: re, im
    end type MyComplex
    interface operator(+)
        module procedure add_complex
    end interface
contains
    function add_complex(a, b) result(c)
        type(MyComplex), intent(in) :: a, b
        type(MyComplex) :: c
        c%re = a%re + b%re
        c%im = a%im + b%im
    end function add_complex
end module complex_mod

program test
    use complex_mod
    type(MyComplex) :: a, b, c
    a = MyComplex(1.0, 2.0)
    b = MyComplex(3.0, 4.0)
    c = a + b
    print *, c%re
end program test
"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn operator_multiply() {
    let out = run_prints(
        r#"
module vec_mod
    implicit none
    type :: Vec3
        real :: x, y, z
    end type Vec3
    interface operator(*)
        module procedure scale_vec
    end interface
contains
    function scale_vec(s, v) result(r)
        real, intent(in) :: s
        type(Vec3), intent(in) :: v
        type(Vec3) :: r
        r%x = s * v%x; r%y = s * v%y; r%z = s * v%z
    end function scale_vec
end module vec_mod

program test
    use vec_mod
    type(Vec3) :: v, r
    v = Vec3(1.0, 2.0, 3.0)
    r = 2.0 * v
    print *, r%x
end program test
"#,
    );
    assert_eq!(out, vec!["2"]);
}

// ── Generic interfaces ────────────────────────────────────────

#[test]
fn generic_interface() {
    let out = run_prints(
        r#"
module generic_mod
    implicit none
    interface my_abs
        module procedure abs_int, abs_real
    end interface
contains
    function abs_int(x) result(r)
        integer, intent(in) :: x
        integer :: r
        r = abs(x)
    end function abs_int
    function abs_real(x) result(r)
        real, intent(in) :: x
        real :: r
        r = abs(x)
    end function abs_real
end module generic_mod

program test
    use generic_mod
    print *, my_abs(-5)
    print *, int(my_abs(-3.14))
end program test
"#,
    );
    assert_eq!(out, vec!["5", "3"]);
}

// ── ASSOCIATE construct ────────────────────────────────────────

#[test]
fn associate_basic() {
    compile_ok(
        r#"
program test
    type :: Point
        real :: x, y
    end type Point
    type(Point) :: p
    p%x = 3.0
    p%y = 4.0
    associate(xx => p%x, yy => p%y)
        print *, sqrt(xx*xx + yy*yy)
    end associate
end program test
"#,
    );
}

#[test]
fn associate_array_elem() {
    compile_ok(
        r#"
program test
    integer :: a(5) = [10, 20, 30, 40, 50]
    associate(mid => a(3))
        print *, mid
    end associate
end program test
"#,
    );
}

#[test]
fn associate_expr() {
    compile_ok(
        r#"
program test
    real :: x = 3.0, y = 4.0
    associate(hyp => sqrt(x*x + y*y))
        print *, hyp
    end associate
end program test
"#,
    );
}
