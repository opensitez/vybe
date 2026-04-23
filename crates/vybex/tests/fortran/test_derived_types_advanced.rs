use super::helpers::compile_ok;

// ── Constructor syntax ────────────────────────────────────────

#[test]
fn type_constructor_positional() {
    compile_ok(r#"
program test
    type :: Point
        real :: x, y
    end type Point
    type(Point) :: p
    p = Point(3.0, 4.0)
    print *, p%x
end program test
"#);
}

#[test]
fn type_constructor_keyword() {
    compile_ok(r#"
program test
    type :: Color
        integer :: r, g, b
    end type Color
    type(Color) :: red
    red = Color(r=255, g=0, b=0)
    print *, red%r
end program test
"#);
}

#[test]
fn type_default_init() {
    compile_ok(r#"
program test
    type :: Config
        integer :: timeout = 30
        real :: threshold = 0.01
        logical :: debug = .false.
    end type Config
    type(Config) :: cfg
    print *, cfg%timeout
end program test
"#);
}

// ── Nested derived types ──────────────────────────────────────

#[test]
fn nested_type() {
    compile_ok(r#"
program test
    type :: Point
        real :: x, y
    end type Point
    type :: Circle
        type(Point) :: center
        real :: radius
    end type Circle
    type(Circle) :: c
    c%center%x = 0.0
    c%center%y = 0.0
    c%radius = 5.0
    print *, c%radius
end program test
"#);
}

#[test]
fn nested_type_three_deep() {
    compile_ok(r#"
program test
    type :: Coord
        real :: val
    end type Coord
    type :: Point
        type(Coord) :: x, y
    end type Point
    type :: Segment
        type(Point) :: start, finish
    end type Segment
    type(Segment) :: s
    s%start%x%val = 1.0
    print *, s%start%x%val
end program test
"#);
}

// ── Type-bound procedures ─────────────────────────────────────

#[test]
fn type_bound_subroutine() {
    compile_ok(r#"
program test
    type :: Counter
        integer :: n = 0
    contains
        procedure :: inc
        procedure :: get
    end type Counter
    type(Counter) :: c
    call c%inc()
    call c%inc()
    print *, c%get()
contains
    subroutine inc(self)
        class(Counter), intent(inout) :: self
        self%n = self%n + 1
    end subroutine inc
    function get(self) result(v)
        class(Counter), intent(in) :: self
        integer :: v
        v = self%n
    end function get
end program test
"#);
}

#[test]
fn type_bound_function() {
    compile_ok(r#"
program test
    type :: Vector
        real :: x, y
    contains
        procedure :: magnitude
    end type Vector
    type(Vector) :: v
    v%x = 3.0
    v%y = 4.0
    print *, v%magnitude()
contains
    function magnitude(self) result(m)
        class(Vector), intent(in) :: self
        real :: m
        m = sqrt(self%x**2 + self%y**2)
    end function magnitude
end program test
"#);
}

#[test]
fn type_bound_final() {
    compile_ok(r#"
program test
    type :: Resource
        integer :: id
    contains
        final :: cleanup
    end type Resource
contains
    subroutine cleanup(self)
        type(Resource), intent(inout) :: self
        self%id = 0
    end subroutine cleanup
end program test
"#);
}

// ── Polymorphism / CLASS ──────────────────────────────────────

#[test]
fn class_polymorphic_arg() {
    compile_ok(r#"
program test
    type :: Animal
        character(len=20) :: name
    end type Animal
    type, extends(Animal) :: Dog
        character(len=10) :: breed
    end type Dog
    type(Dog) :: d
    d%name = 'Rex'
    d%breed = 'Labrador'
    call describe(d)
contains
    subroutine describe(a)
        class(Animal), intent(in) :: a
        print *, trim(a%name)
    end subroutine describe
end program test
"#);
}

#[test]
fn class_unlimited_polymorphic() {
    compile_ok(r#"
program test
    class(*), pointer :: p => null()
    print *, "ok"
end program test
"#);
}

#[test]
fn select_type_basic() {
    compile_ok(r#"
program test
    type :: Base
        integer :: id = 0
    end type Base
    type, extends(Base) :: Child
        integer :: extra = 99
    end type Child
    class(Base), allocatable :: obj
    allocate(Child :: obj)
    select type(obj)
    type is (Child)
        print *, obj%extra
    class default
        print *, "base"
    end select
end program test
"#);
}

// ── SEQUENCE attribute ─────────────────────────────────────────

#[test]
fn sequence_type() {
    compile_ok(r#"
program test
    type :: Packed
        sequence
        integer :: a
        real :: b
        logical :: c
    end type Packed
    type(Packed) :: p
    p%a = 1
    p%b = 2.0
    p%c = .true.
    print *, p%a
end program test
"#);
}

// ── Array of derived types ────────────────────────────────────

#[test]
fn array_of_types() {
    compile_ok(r#"
program test
    type :: Point
        real :: x, y
    end type Point
    type(Point) :: pts(3)
    integer :: i
    do i = 1, 3
        pts(i)%x = real(i)
        pts(i)%y = real(i) * 2.0
    end do
    print *, pts(2)%x
end program test
"#);
}

#[test]
fn allocatable_type_array() {
    compile_ok(r#"
program test
    type :: Node
        integer :: value
    end type Node
    type(Node), allocatable :: nodes(:)
    allocate(nodes(5))
    nodes(1)%value = 42
    print *, nodes(1)%value
    deallocate(nodes)
end program test
"#);
}

// ── Type comparison and assignment ────────────────────────────

#[test]
fn type_assignment() {
    compile_ok(r#"
program test
    type :: Box
        integer :: width, height
    end type Box
    type(Box) :: a, b
    a%width = 10
    a%height = 20
    b = a
    print *, b%width
end program test
"#);
}

// ── Type in module, used in program ──────────────────────────

#[test]
fn module_type_export_use() {
    compile_ok(r#"
module shapes
    implicit none
    type :: Rectangle
        real :: width, height
    end type Rectangle
contains
    function area(r) result(a)
        type(Rectangle), intent(in) :: r
        real :: a
        a = r%width * r%height
    end function area
end module shapes

program test
    use shapes
    type(Rectangle) :: rect
    rect%width = 5.0
    rect%height = 3.0
    print *, area(rect)
end program test
"#);
}
