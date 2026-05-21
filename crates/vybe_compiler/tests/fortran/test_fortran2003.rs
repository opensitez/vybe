use super::helpers::{compile_ok, run_prints};
use vybe_compiler::ast::{ClassMember, StmtKind, Visibility};

// ── Abstract types and deferred procedures ────────────────────

#[test] fn abstract_type_basic() {
    compile_ok(r#"
program test
    type, abstract :: Shape
        real :: color(3)
    contains
        procedure(area_iface), deferred :: area
    end type Shape
    print *, "ok"
end program test

abstract interface
    function area_iface(self) result(a)
        import Shape
        class(Shape), intent(in) :: self
        real :: a
    end function area_iface
end interface
"#);
}

#[test]
fn abstract_type_attributes_are_preserved_in_ast() {
    let module = vybe_compiler::languages::fortran::parse(r#"
module shapes
    implicit none

    type :: Base
    end type Base

    type, abstract, extends(Base), private :: Shape
    contains
        procedure(area_iface), deferred, private, non_overridable :: area
    end type Shape

    abstract interface
        function area_iface(self) result(a)
            import Shape
            class(Shape), intent(in) :: self
            real :: a
        end function area_iface
    end interface
end module shapes
"#).expect("parse failed");

    let (parents, modifiers, members) = module
        .body
        .iter()
        .find_map(|statement| match &statement.kind {
            StmtKind::ModuleDecl { members, .. } => members.iter().find_map(|member| match member {
                ClassMember::NestedType(stmt) => match &stmt.kind {
                    StmtKind::ClassDecl {
                        name,
                        parents,
                        modifiers,
                        members,
                        ..
                    } if name.eq_ignore_ascii_case("Shape") => Some((parents, modifiers, members)),
                    _ => None,
                },
                _ => None,
            }),
            _ => None,
        })
        .expect("missing Shape declaration");

    assert_eq!(parents, &["Base".to_string()]);
    assert_eq!(modifiers.visibility, Visibility::Private);
    assert!(modifiers.is_abstract);

    let method_names: Vec<&str> = members
        .iter()
        .filter_map(|member| match member {
            ClassMember::Method(stmt) => match &stmt.kind {
                StmtKind::FunctionDecl { name, .. } => Some(name.as_str()),
                _ => None,
            },
            _ => None,
        })
        .collect();
    assert!(method_names.contains(&"area"));
    assert!(!method_names.contains(&"area_iface"));

    let area_modifiers = members
        .iter()
        .find_map(|member| match member {
            ClassMember::Method(stmt) => match &stmt.kind {
                StmtKind::FunctionDecl { name, modifiers, .. } if name.eq_ignore_ascii_case("area") => {
                    Some(modifiers)
                }
                _ => None,
            },
            _ => None,
        })
        .expect("missing area binding");

    assert_eq!(area_modifiers.visibility, Visibility::Private);
    assert!(area_modifiers.is_abstract);
    assert!(area_modifiers.is_not_overridable);
}

#[test] fn abstract_type_extended() {
    compile_ok(r#"
module shapes
    implicit none
    type, abstract :: Shape
    contains
        procedure(compute_area), deferred :: area
    end type Shape

    abstract interface
        function compute_area(self) result(a)
            import Shape
            class(Shape), intent(in) :: self
            real :: a
        end function compute_area
    end interface

    type, extends(Shape) :: Circle
        real :: radius
    contains
        procedure :: area => circle_area
    end type Circle

contains
    function circle_area(self) result(a)
        class(Circle), intent(in) :: self
        real :: a
        a = 3.14159 * self%radius ** 2
    end function circle_area
end module shapes

program test
    use shapes
    type(Circle) :: c
    c%radius = 5.0
    print *, c%area()
end program test
"#);
}

#[test] fn type_bound_procedure_allows_keyword_binding_name() {
    let out = run_prints(r#"
program test
    type :: stats_result
        integer :: n = 12
    contains
        procedure :: print => print_stats
    end type stats_result

    type(stats_result) :: stats
    call stats%print()
contains
    subroutine print_stats(self)
        class(stats_result), intent(in) :: self
        print *, "n =", self%n
    end subroutine print_stats
end program test
"#);
    assert_eq!(out, ["n = 12"]);
}

#[test]
fn type_bound_generic_binding_alias_runs() {
    let out = run_prints(r#"
program test
    type :: Counter
        integer :: n = 4
    contains
        procedure :: doubled_impl
        generic :: doubled => doubled_impl
    end type Counter

    type(Counter) :: value
    print *, value%doubled()
contains
    integer function doubled_impl(self) result(v)
        class(Counter), intent(in) :: self
        integer :: v
        v = self%n * 2
    end function doubled_impl
end program test
"#);

    assert_eq!(out, ["8"]);
}

#[test] fn deferred_binding() {
    compile_ok(r#"
module iface_mod
    implicit none
    type, abstract :: Base
    contains
        procedure(greet_iface), deferred :: greet
    end type Base

    abstract interface
        subroutine greet_iface(self)
            import Base
            class(Base), intent(in) :: self
        end subroutine greet_iface
    end interface
end module iface_mod

program test
    print *, "ok"
end program test
"#);
}

// ── CLASS(*) — unlimited polymorphism ────────────────────────

#[test] fn class_star_pointer() {
    compile_ok(r#"
program test
    class(*), pointer :: p => null()
    integer, target :: x = 42
    p => x
    print *, "ok"
end program test
"#);
}

#[test] fn class_star_allocatable() {
    compile_ok(r#"
program test
    class(*), allocatable :: obj
    allocate(integer :: obj)
    print *, "ok"
end program test
"#);
}

#[test] fn select_type_unlimited() {
    compile_ok(r#"
program test
    class(*), allocatable :: val
    allocate(integer :: val)
    select type(val)
    type is (integer)
        print *, 'integer'
    type is (real)
        print *, 'real'
    class default
        print *, 'other'
    end select
end program test
"#);
}

// ── Polymorphic arguments ─────────────────────────────────────

#[test] fn polymorphic_arg_in() {
    compile_ok(r#"
program test
    type :: Animal
        character(len=20) :: name = 'unknown'
    end type Animal
    type, extends(Animal) :: Dog
    end type Dog
    type(Dog) :: d
    d%name = 'Rex'
    call show_name(d)
contains
    subroutine show_name(a)
        class(Animal), intent(in) :: a
        print *, trim(a%name)
    end subroutine show_name
end program test
"#);
}

#[test] fn polymorphic_allocatable() {
    compile_ok(r#"
program test
    type :: Vehicle
        integer :: wheels = 4
    end type Vehicle
    type, extends(Vehicle) :: Bike
    end type Bike
    class(Vehicle), allocatable :: v
    allocate(Bike :: v)
    print *, v%wheels
end program test
"#);
}

#[test] fn polymorphic_array() {
    compile_ok(r#"
program test
    type :: Base
        integer :: id = 0
    end type Base
    class(Base), allocatable :: arr(:)
    allocate(Base :: arr(3))
    arr(1)%id = 1
    print *, arr(1)%id
end program test
"#);
}

// ── TYPE_IS vs CLASS_IS in SELECT TYPE ───────────────────────

#[test] fn select_type_class_is() {
    compile_ok(r#"
program test
    type :: A
        integer :: x = 1
    end type A
    type, extends(A) :: B
        integer :: y = 2
    end type B
    class(A), allocatable :: obj
    allocate(B :: obj)
    select type(obj)
    class is (B)
        print *, obj%y
    type is (A)
        print *, obj%x
    end select
end program test
"#);
}

// ── ENUM (Fortran 2003) ───────────────────────────────────────

#[test] fn enum_basic() {
    compile_ok(r#"
program test
    enum, bind(c)
        enumerator :: RED = 0, GREEN = 1, BLUE = 2
    end enum
    integer :: color = GREEN
    print *, color
end program test
"#);
}

#[test] fn enum_auto_increment() {
    compile_ok(r#"
program test
    enum, bind(c)
        enumerator :: NORTH, SOUTH, EAST, WEST
    end enum
    integer :: dir = EAST
    print *, dir
end program test
"#);
}

#[test] fn enum_explicit_values() {
    compile_ok(r#"
program test
    enum, bind(c)
        enumerator :: LOW = 1, MEDIUM = 5, HIGH = 10
    end enum
    integer :: level = HIGH
    print *, level
end program test
"#);
}

// ── C Interoperability (iso_c_binding) ────────────────────────

#[test] fn iso_c_binding_use() {
    compile_ok(r#"
program test
    use iso_c_binding
    integer(c_int) :: n = 42_c_int
    real(c_double) :: x = 3.14_c_double
    print *, n
    print *, x
end program test
"#);
}

#[test] fn bind_c_function() {
    compile_ok(r#"
module c_funcs
    use iso_c_binding
    implicit none
    interface
        function c_strlen(s) bind(c, name='strlen') result(n)
            use iso_c_binding
            type(c_ptr), value :: s
            integer(c_size_t) :: n
        end function c_strlen
    end interface
end module c_funcs

program test
    print *, "ok"
end program test
"#);
}

#[test] fn bind_c_type() {
    compile_ok(r#"
program test
    use iso_c_binding
    type, bind(c) :: point_t
        real(c_float) :: x
        real(c_float) :: y
    end type point_t
    type(point_t) :: p
    p%x = 1.0
    p%y = 2.0
    print *, p%x
end program test
"#);
}

#[test] fn c_interop_kinds() {
    compile_ok(r#"
program test
    use iso_c_binding
    integer(c_int)    :: i = 1_c_int
    integer(c_long)   :: j = 2_c_long
    integer(c_size_t) :: k = 3_c_size_t
    real(c_float)     :: f = 1.0_c_float
    real(c_double)    :: d = 2.0_c_double
    logical(c_bool)   :: b = .true._c_bool
    character(len=1, kind=c_char) :: ch = c_null_char
    print *, i, j, f
end program test
"#);
}

// ── PROTECTED attribute ───────────────────────────────────────

#[test] fn protected_variable() {
    compile_ok(r#"
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
"#);
}

// ── VOLATILE ─────────────────────────────────────────────────

#[test] fn volatile_integer() {
    compile_ok(r#"
program test
    integer, volatile :: x = 0
    x = 42
    print *, x
end program test
"#);
}

#[test] fn volatile_in_module() {
    compile_ok(r#"
module hw_reg
    implicit none
    integer, volatile :: status_reg = 0
    integer, volatile :: data_reg = 0
end module hw_reg

program test
    use hw_reg
    status_reg = 1
    print *, status_reg
end program test
"#);
}

// ── ISO_FORTRAN_ENV ───────────────────────────────────────────

#[test] fn iso_fortran_env_kinds() {
    compile_ok(r#"
program test
    use iso_fortran_env
    integer(int32) :: n = 42_int32
    integer(int64) :: big = 1000000000_int64
    real(real32) :: f = 3.14_real32
    real(real64) :: d = 3.14159265_real64
    print *, n
    print *, big
end program test
"#);
}

#[test] fn iso_fortran_env_units() {
    compile_ok(r#"
program test
    use iso_fortran_env
    write(output_unit, *) 'stdout'
    write(error_unit, *) 'stderr'
end program test
"#);
}

#[test] fn iso_fortran_env_compiler_version() {
    compile_ok(r#"
program test
    use iso_fortran_env
    print *, compiler_version()
    print *, compiler_options()
end program test
"#);
}

// ── MOVE_ALLOC (Fortran 2003) ─────────────────────────────────

#[test] fn move_alloc_basic() {
    compile_ok(r#"
program test
    integer, allocatable :: a(:), b(:)
    allocate(a(3))
    a = [1, 2, 3]
    call move_alloc(a, b)
    print *, b(1)
    print *, allocated(a)
end program test
"#);
}

// ── ALLOCATED intrinsic ───────────────────────────────────────

#[test] fn allocated_false() {
    compile_ok(r#"
program test
    integer, allocatable :: x(:)
    print *, allocated(x)
end program test
"#);
}

#[test] fn allocated_true() {
    compile_ok(r#"
program test
    integer, allocatable :: x(:)
    allocate(x(5))
    print *, allocated(x)
    deallocate(x)
end program test
"#);
}

// ── EXTENDS_TYPE_OF and SAME_TYPE_AS ─────────────────────────

#[test] fn same_type_as() {
    compile_ok(r#"
program test
    type :: A
        integer :: x = 1
    end type A
    type(A) :: obj1, obj2
    print *, same_type_as(obj1, obj2)
end program test
"#);
}

#[test] fn extends_type_of() {
    compile_ok(r#"
program test
    type :: Base
        integer :: x = 0
    end type Base
    type, extends(Base) :: Child
        integer :: y = 1
    end type Child
    type(Base) :: b
    type(Child) :: c
    print *, extends_type_of(c, b)
end program test
"#);
}
