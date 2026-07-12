use super::helpers::{compile_ok, run_prints};

// ── Constructor syntax ────────────────────────────────────────

#[test]
fn type_constructor_positional() {
    let out = run_prints(
        r#"
program test
    type :: Point
        real :: x, y
    end type Point
    type(Point) :: p
    p = Point(3.0, 4.0)
    print *, p%x
end program test
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn type_constructor_keyword() {
    let out = run_prints(
        r#"
program test
    type :: Color
        integer :: r, g, b
    end type Color
    type(Color) :: red
    red = Color(r=255, g=0, b=0)
    print *, red%r
end program test
"#,
    );
    assert_eq!(out, vec!["255"]);
}

#[test]
fn type_default_init() {
    compile_ok(
        r#"
program test
    type :: Config
        integer :: timeout = 30
        real :: threshold = 0.01
        logical :: debug = .false.
    end type Config
    type(Config) :: cfg
    print *, cfg%timeout
end program test
"#,
    );
}

// ── Nested derived types ──────────────────────────────────────

#[test]
fn nested_type() {
    compile_ok(
        r#"
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
"#,
    );
}

#[test]
fn nested_type_three_deep() {
    compile_ok(
        r#"
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
"#,
    );
}

// ── Type-bound procedures ─────────────────────────────────────

#[test]
fn type_bound_subroutine() {
    let out = run_prints(
        r#"
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
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn module_type_bound_subroutine() {
    let out = run_prints(
        r#"
module counters
    implicit none

    type :: Counter
        integer :: n = 0
    contains
        procedure :: inc
        procedure :: get
    end type Counter
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
end module counters

program test
    use counters
    implicit none
    type(Counter) :: c
    call c%inc()
    call c%inc()
    print *, c%get()
end program test
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn subroutine_populates_derived_type_out_param() {
    let out = run_prints(
        r#"
program test
    type :: Counter
        integer :: n = 0
    end type Counter
    type(Counter) :: c
    call fill(c)
    print *, c%n
contains
    subroutine fill(counter)
        type(Counter), intent(out) :: counter
        counter%n = 7
    end subroutine fill
end program test
"#,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn assumed_shape_intrinsics_populate_derived_type_out_param() {
    let out = run_prints(
        r#"
program test
    type :: Stats
        integer :: n = 0
        real :: lo = 0.0
        real :: hi = 0.0
    end type Stats
    real :: a(3) = [1.0, 2.0, 3.0]
    type(Stats) :: s
    call fill(a, s)
    print *, s%n
    print *, s%lo
    print *, s%hi
contains
    subroutine fill(data, result)
        real, intent(in) :: data(:)
        type(Stats), intent(out) :: result
        result%n = size(data)
        result%lo = minval(data)
        result%hi = maxval(data)
    end subroutine fill
end program test
"#,
    );
    assert_eq!(out, vec!["3", "1", "3"]);
}

#[test]
fn type_bound_function() {
    compile_ok(
        r#"
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
"#,
    );
}

#[test]
fn type_bound_final() {
    compile_ok(
        r#"
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
"#,
    );
}

// ── Polymorphism / CLASS ──────────────────────────────────────

#[test]
fn class_polymorphic_arg() {
    compile_ok(
        r#"
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
"#,
    );
}

#[test]
fn class_unlimited_polymorphic() {
    compile_ok(
        r#"
program test
    class(*), pointer :: p => null()
    print *, "ok"
end program test
"#,
    );
}

#[test]
fn select_type_basic() {
    compile_ok(
        r#"
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
"#,
    );
}

// ── SEQUENCE attribute ─────────────────────────────────────────

#[test]
fn sequence_type() {
    compile_ok(
        r#"
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
"#,
    );
}

// ── Array of derived types ────────────────────────────────────

#[test]
fn array_of_types() {
    compile_ok(
        r#"
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
"#,
    );
}

#[test]
fn allocatable_type_array() {
    let out = run_prints(
        r#"
program test
    type :: Node
        integer :: value
    end type Node
    type(Node), allocatable :: nodes(:)
    allocate(nodes(5))
    nodes(1)%value = 42
    print *, nodes(1)%value
    nodes(5)%value = 9
    print *, nodes(5)%value
    deallocate(nodes)
end program test
"#,
    );
    assert_eq!(out, ["42", "9"]);
}

// ── Type comparison and assignment ────────────────────────────

#[test]
fn type_assignment() {
    compile_ok(
        r#"
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
"#,
    );
}

// ── Type in module, used in program ──────────────────────────

#[test]
fn module_type_export_use() {
    compile_ok(
        r#"
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
"#,
    );
}

#[test]
fn tbp_with_print_method_name() {
    let out = run_prints(
        r#"
module swe_types
    implicit none
    type :: grid_t
        integer  :: nx, ny
    contains
        procedure :: init  => grid_init
        procedure :: print => grid_print
    end type grid_t
contains
    subroutine grid_init(self, nx, ny)
        class(grid_t), intent(inout) :: self
        integer, intent(in) :: nx, ny
        self%nx = nx
        self%ny = ny
    end subroutine grid_init
    subroutine grid_print(self)
        class(grid_t), intent(in) :: self
        print *, self%nx
        print *, self%ny
    end subroutine grid_print
end module swe_types
program test_weather
    use swe_types
    type(grid_t) :: grid
    call grid%init(10, 20)
    call grid%print()
end program test_weather
"#,
    );
    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn nested_derived_types_with_tbp() {
    let out = run_prints(
        r#"
module swe_types
    implicit none
    type :: field2d
        integer :: nx, ny
        character(len=32) :: name
    contains
        procedure :: init    => field_init
    end type field2d
    type :: swe_state
        type(field2d) :: h
        type(field2d) :: u
    end type swe_state
contains
    subroutine field_init(self, nx, ny, name)
        class(field2d), intent(inout) :: self
        integer, intent(in) :: nx, ny
        character(len=*), intent(in) :: name
        self%nx = nx
        self%ny = ny
        self%name = name
    end subroutine field_init
end module swe_types
program test
    use swe_types
    type(swe_state) :: state
    call state%h%init(4, 4, "depth")
    call state%u%init(4, 4, "u-vel")
    print *, state%h%nx
    print *, state%h%name
    print *, state%u%name
end program test
"#,
    );
    assert_eq!(out, vec!["4", "depth", "u-vel"]);
}

#[test]
fn weather_model_grid_print_tbp() {
    // Direct reproduction of the weather model's grid_t with print TBP
    let out = run_prints(
        r#"
module swe_types
    implicit none
    integer, parameter :: dp = kind(1.0d0)
    real(dp), parameter :: G = 9.81_dp

    type :: grid_t
        integer  :: nx, ny
        real(dp) :: dx, dy
        real(dp) :: Lx, Ly
    contains
        procedure :: init  => grid_init
        procedure :: print => grid_print
    end type grid_t

contains
    subroutine grid_init(self, nx, ny, Lx, Ly)
        class(grid_t), intent(inout) :: self
        integer, intent(in) :: nx, ny
        real(dp), intent(in) :: Lx, Ly
        self%nx = nx
        self%ny = ny
        self%Lx = Lx
        self%Ly = Ly
        self%dx = Lx / nx
        self%dy = Ly / ny
    end subroutine grid_init

    subroutine grid_print(self)
        class(grid_t), intent(in) :: self
        print *, self%nx
        print *, self%ny
    end subroutine grid_print
end module swe_types

program test_weather
    use swe_types
    type(grid_t) :: grid
    call grid%init(10, 10, 1000.0d0, 1000.0d0)
    print *, "before print"
    call grid%print()
    print *, "after print"
end program test_weather
"#,
    );
    assert_eq!(out, vec!["before print", "10", "10", "after print"]);
}

#[test]
fn weather_model_full_types() {
    // More complete weather model structure
    let out = run_prints(
        r#"
module swe_types
    implicit none
    integer, parameter :: dp = kind(1.0d0)
    real(dp), parameter :: G = 9.81_dp

    type :: grid_t
        integer  :: nx, ny
        real(dp) :: Lx, Ly
    contains
        procedure :: init  => grid_init
        procedure :: print => grid_print
    end type grid_t

    type :: field2d
        integer :: nx, ny
        character(len=32) :: name
    contains
        procedure :: init    => field_init
    end type field2d

    type :: swe_state
        type(field2d) :: h
        type(field2d) :: u
        type(field2d) :: v
        real(dp)      :: time
    end type swe_state

    type :: swe_config
        integer  :: nx, ny
        integer  :: nt
        real(dp) :: dt
    end type swe_config

contains
    subroutine grid_init(self, nx, ny, Lx, Ly)
        class(grid_t), intent(inout) :: self
        integer, intent(in) :: nx, ny
        real(dp), intent(in) :: Lx, Ly
        self%nx = nx
        self%ny = ny
        self%Lx = Lx
        self%Ly = Ly
    end subroutine grid_init

    subroutine grid_print(self)
        class(grid_t), intent(in) :: self
        print *, self%nx
        print *, self%ny
    end subroutine grid_print

    subroutine field_init(self, nx, ny, name)
        class(field2d), intent(inout) :: self
        integer, intent(in) :: nx, ny
        character(len=*), intent(in) :: name
        self%nx = nx
        self%ny = ny
        self%name = name
    end subroutine field_init
end module swe_types

program test_weather
    use swe_types
    type(grid_t)    :: grid
    type(swe_state) :: state
    type(swe_config) :: cfg

    call grid%init(10, 10, 1000.0d0, 1000.0d0)
    call state%h%init(10, 10, "depth")
    call grid%print()
    print *, state%h%name
end program test_weather
"#,
    );
    assert_eq!(out, vec!["10", "10", "depth"]);
}

// ── Weather model crash reproduction ─────────────────────────────────────────
// Reproduces: call grid%print() fails "undefined is not callable" at runtime.
// examples/fortran/weather_model.f90 line 492.

#[test]
fn tbp_call_after_allocatable_field_init() {
    // grid_t has allocatable array fields; grid_init allocates them;
    // then the TBP `print` must still be callable.
    let out = run_prints(
        r#"
module swe_types
    implicit none
    integer, parameter :: dp = kind(1.0d0)

    type :: grid_t
        integer  :: nx, ny
        real(dp) :: dx, dy, Lx, Ly
        real(dp), allocatable :: x(:), y(:)
    contains
        procedure :: init  => grid_init
        procedure :: print => grid_print
    end type grid_t

contains
    subroutine grid_init(self, nx, ny, Lx, Ly)
        class(grid_t), intent(inout) :: self
        integer,  intent(in) :: nx, ny
        real(dp), intent(in) :: Lx, Ly
        integer :: i
        self%nx = nx;  self%ny = ny
        self%Lx = Lx;  self%Ly = Ly
        self%dx = Lx / nx
        self%dy = Ly / ny
        allocate(self%x(nx), self%y(ny))
        self%x = [(( i - 0.5_dp) * self%dx, i = 1, nx)]
        self%y = [(( i - 0.5_dp) * self%dy, i = 1, ny)]
    end subroutine grid_init

    subroutine grid_print(self)
        class(grid_t), intent(in) :: self
        print *, self%nx
        print *, self%ny
    end subroutine grid_print
end module swe_types

program test
    use swe_types
    type(grid_t) :: grid
    call grid%init(4, 4, 1000.0d0, 1000.0d0)
    call grid%print()
end program test
"#,
    );
    assert_eq!(out, vec!["4", "4"]);
}

#[test]
fn tbp_call_across_multiple_use_modules() {
    // swe_numerics re-uses swe_types; program uses both.
    // Tests that TBP bindings survive multi-module compilation.
    let out = run_prints(
        r#"
module swe_types
    implicit none
    integer, parameter :: dp = kind(1.0d0)

    type :: grid_t
        integer  :: nx, ny
        real(dp) :: Lx, Ly, dx, dy
        real(dp), allocatable :: x(:), y(:)
    contains
        procedure :: init  => grid_init
        procedure :: print => grid_print
    end type grid_t

contains
    subroutine grid_init(self, nx, ny, Lx, Ly)
        class(grid_t), intent(inout) :: self
        integer,  intent(in) :: nx, ny
        real(dp), intent(in) :: Lx, Ly
        integer :: i
        self%nx = nx;  self%ny = ny
        self%Lx = Lx;  self%Ly = Ly
        self%dx = Lx / nx
        self%dy = Ly / ny
        allocate(self%x(nx), self%y(ny))
        self%x = [(( i - 0.5_dp) * self%dx, i = 1, nx)]
        self%y = [(( i - 0.5_dp) * self%dy, i = 1, ny)]
    end subroutine grid_init

    subroutine grid_print(self)
        class(grid_t), intent(in) :: self
        print *, self%nx
    end subroutine grid_print
end module swe_types

module swe_numerics
    use swe_types
    implicit none
contains
    pure function wrap(i, n) result(j)
        integer, intent(in) :: i, n
        integer :: j
        j = mod(i - 1 + n, n) + 1
    end function wrap
end module swe_numerics

program test
    use swe_types
    use swe_numerics
    type(grid_t) :: grid
    call grid%init(8, 8, 2000.0d0, 2000.0d0)
    call grid%print()
end program test
"#,
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn weather_model_full_reproduction() {
    // Full weather model setup: three modules, allocatable fields,
    // array constructors, logical restart path, then grid%print().
    let out = run_prints(
        r#"
module swe_types
    implicit none
    integer, parameter :: dp = kind(1.0d0)
    real(dp), parameter :: G = 9.81_dp

    type :: grid_t
        integer  :: nx, ny
        real(dp) :: dx, dy, Lx, Ly
        real(dp), allocatable :: x(:), y(:)
    contains
        procedure :: init  => grid_init
        procedure :: print => grid_print
    end type grid_t

    type :: field2d
        real(dp), allocatable :: data(:,:)
        integer :: nx, ny
        character(len=32) :: name
    contains
        procedure :: init => field_init
    end type field2d

    type :: swe_state
        type(field2d) :: h
        type(field2d) :: u
        type(field2d) :: v
        real(dp) :: time
    end type swe_state

    type :: swe_config
        integer  :: nx, ny
        real(dp) :: Lx, Ly, dt, f_coriolis, h0
    end type swe_config

contains
    subroutine grid_init(self, nx, ny, Lx, Ly)
        class(grid_t), intent(inout) :: self
        integer,  intent(in) :: nx, ny
        real(dp), intent(in) :: Lx, Ly
        integer :: i
        self%nx = nx;  self%ny = ny
        self%Lx = Lx;  self%Ly = Ly
        self%dx = Lx / nx
        self%dy = Ly / ny
        allocate(self%x(nx), self%y(ny))
        self%x = [(( i - 0.5_dp) * self%dx, i = 1, nx)]
        self%y = [(( i - 0.5_dp) * self%dy, i = 1, ny)]
    end subroutine grid_init

    subroutine grid_print(self)
        class(grid_t), intent(in) :: self
        print "(a, i0, a, i0)", "Grid: ", self%nx, " x ", self%ny
    end subroutine grid_print

    subroutine field_init(self, nx, ny, name)
        class(field2d), intent(inout) :: self
        integer, intent(in) :: nx, ny
        character(len=*), intent(in) :: name
        self%nx = nx;  self%ny = ny
        self%name = name
        allocate(self%data(nx, ny))
        self%data = 0.0_dp
    end subroutine field_init
end module swe_types

module swe_numerics
    use swe_types
    implicit none
contains
    pure function wrap(i, n) result(j)
        integer, intent(in) :: i, n
        integer :: j
        j = mod(i - 1 + n, n) + 1
    end function wrap
end module swe_numerics

program weather_model
    use swe_types
    use swe_numerics
    implicit none

    type(swe_state)  :: state
    type(grid_t)     :: grid
    type(swe_config) :: cfg
    logical :: restart_ok

    cfg%nx = 4;  cfg%ny = 4
    cfg%Lx = 1.0e6_dp;  cfg%Ly = 1.0e6_dp
    cfg%dt = 60.0_dp
    cfg%f_coriolis = 1.0e-4_dp
    cfg%h0 = 1000.0_dp

    call grid%init(cfg%nx, cfg%ny, cfg%Lx, cfg%Ly)
    call state%h%init(cfg%nx, cfg%ny, "h")
    call state%u%init(cfg%nx, cfg%ny, "u")
    call state%v%init(cfg%nx, cfg%ny, "v")
    state%time = 0.0_dp

    restart_ok = .false.

    if (.not. restart_ok) then
        print *, "Initialized vortex"
    end if

    print *, "Header"
    call grid%print()
    print *, "Done"
end program weather_model
"#,
    );
    assert_eq!(
        out,
        vec!["Initialized vortex", "Header", "Grid: 4 x 4", "Done"]
    );
}

#[test]
fn tbp_call_after_nested_do_loop_data_write() {
    // The weather model crashes after nested do-loops write into field2d%data(i,j).
    // Test that grid%print() TBP works after that pattern.
    let out = run_prints(
        r#"
module swe_types
    implicit none
    integer, parameter :: dp = kind(1.0d0)
    real(dp), parameter :: G = 9.81_dp

    type :: grid_t
        integer  :: nx, ny
        real(dp) :: dx, dy, Lx, Ly
        real(dp), allocatable :: x(:), y(:)
    contains
        procedure :: init  => grid_init
        procedure :: print => grid_print
    end type grid_t

    type :: field2d
        real(dp), allocatable :: data(:,:)
        integer :: nx, ny
        character(len=32) :: name
    contains
        procedure :: init => field_init
    end type field2d

    type :: swe_state
        type(field2d) :: h
        type(field2d) :: u
        type(field2d) :: v
        real(dp) :: time
    end type swe_state

contains
    subroutine grid_init(self, nx, ny, Lx, Ly)
        class(grid_t), intent(inout) :: self
        integer,  intent(in) :: nx, ny
        real(dp), intent(in) :: Lx, Ly
        integer :: i
        self%nx = nx;  self%ny = ny
        self%Lx = Lx;  self%Ly = Ly
        self%dx = Lx / nx
        self%dy = Ly / ny
        allocate(self%x(nx), self%y(ny))
        self%x = [(( i - 0.5_dp) * self%dx, i = 1, nx)]
        self%y = [(( i - 0.5_dp) * self%dy, i = 1, ny)]
    end subroutine grid_init

    subroutine grid_print(self)
        class(grid_t), intent(in) :: self
        print *, self%nx
        print *, self%ny
    end subroutine grid_print

    subroutine field_init(self, nx, ny, name)
        class(field2d), intent(inout) :: self
        integer, intent(in) :: nx, ny
        character(len=*), intent(in) :: name
        self%nx = nx;  self%ny = ny
        self%name = name
        allocate(self%data(nx, ny))
        self%data = 0.0_dp
    end subroutine field_init
end module swe_types

program test
    use swe_types
    implicit none
    type(swe_state) :: state
    type(grid_t)    :: grid
    integer :: i, j
    real(dp) :: x, y, r2, amp
    integer, parameter :: nx = 4, ny = 4
    real(dp), parameter :: Lx = 1.0e6_dp, Ly = 1.0e6_dp
    real(dp), parameter :: f = 1.0e-4_dp, h0 = 1000.0_dp

    call grid%init(nx, ny, Lx, Ly)
    call state%h%init(nx, ny, "h")
    call state%u%init(nx, ny, "u")
    call state%v%init(nx, ny, "v")
    state%time = 0.0_dp

    ! Geostrophic vortex initial condition (the nested do-loops from weather_model)
    amp = 50.0_dp
    do j = 1, ny
        do i = 1, nx
            x  = grid%x(i) - Lx * 0.5_dp
            y  = grid%y(j) - Ly * 0.5_dp
            r2 = (x**2 + y**2) / (2.0e5_dp)**2

            state%h%data(i,j) = h0 + amp * exp(-r2)

            state%u%data(i,j) = -(G * amp / f) * &
                (-2.0_dp * y / (2.0e5_dp)**2) * exp(-r2)
            state%v%data(i,j) =  (G * amp / f) * &
                (-2.0_dp * x / (2.0e5_dp)**2) * exp(-r2)
        end do
    end do

    print *, "vortex done"
    call grid%print()
    print *, "print done"
end program test
"#,
    );
    assert_eq!(out, vec!["vortex done", "4", "4", "print done"]);
}

#[test]
fn tbp_call_with_swe_io_module() {
    // Three-module structure like weather_model.f90 (swe_types + swe_numerics + swe_io).
    // grid%print() must work after read_restart returns ok=.false. on a
    // deterministic negative path.
    let out = run_prints(
        r#"
module swe_types
    implicit none
    integer, parameter :: dp = kind(1.0d0)

    type :: grid_t
        integer  :: nx, ny
        real(dp) :: dx, dy, Lx, Ly
        real(dp), allocatable :: x(:), y(:)
    contains
        procedure :: init  => grid_init
        procedure :: print => grid_print
    end type grid_t

    type :: field2d
        real(dp), allocatable :: data(:,:)
        integer :: nx, ny
        character(len=32) :: name
    contains
        procedure :: init => field_init
    end type field2d

    type :: swe_state
        type(field2d) :: h
        type(field2d) :: u
        type(field2d) :: v
        real(dp) :: time
    end type swe_state

contains
    subroutine grid_init(self, nx, ny, Lx, Ly)
        class(grid_t), intent(inout) :: self
        integer,  intent(in) :: nx, ny
        real(dp), intent(in) :: Lx, Ly
        integer :: i
        self%nx = nx;  self%ny = ny
        self%Lx = Lx;  self%Ly = Ly
        self%dx = Lx / nx
        self%dy = Ly / ny
        allocate(self%x(nx), self%y(ny))
        self%x = [(( i - 0.5_dp) * self%dx, i = 1, nx)]
        self%y = [(( i - 0.5_dp) * self%dy, i = 1, ny)]
    end subroutine grid_init

    subroutine grid_print(self)
        class(grid_t), intent(in) :: self
        print *, self%nx
    end subroutine grid_print

    subroutine field_init(self, nx, ny, name)
        class(field2d), intent(inout) :: self
        integer, intent(in) :: nx, ny
        character(len=*), intent(in) :: name
        self%nx = nx;  self%ny = ny
        self%name = name
        allocate(self%data(nx, ny))
        self%data = 0.0_dp
    end subroutine field_init
end module swe_types

module swe_numerics
    use swe_types
    implicit none
contains
    pure function wrap(i, n) result(j)
        integer, intent(in) :: i, n
        integer :: j
        j = mod(i - 1 + n, n) + 1
    end function wrap
end module swe_numerics

module swe_io
    use swe_types
    implicit none
contains
    subroutine write_restart(state, grid, filename)
        type(swe_state), intent(in) :: state
        type(grid_t),    intent(in) :: grid
        character(len=*), intent(in) :: filename
        integer :: unit
        open(newunit=unit, file=trim(filename), form="unformatted", &
             status="replace", action="write")
        write(unit) state%time
        write(unit) grid%nx, grid%ny
        write(unit) state%h%data
        write(unit) state%u%data
        write(unit) state%v%data
        close(unit)
    end subroutine write_restart

    subroutine read_restart(state, grid, filename, ok)
        type(swe_state), intent(inout) :: state
        type(grid_t),    intent(in)    :: grid
        character(len=*), intent(in)   :: filename
        logical, intent(out) :: ok
        integer :: unit, nx, ny, ios
        open(newunit=unit, file=trim(filename), form="unformatted", &
             status="old", action="read", iostat=ios)
        if (ios /= 0) then
            ok = .false.
            return
        end if
        read(unit) state%time
        read(unit) nx, ny
        if (nx /= grid%nx .or. ny /= grid%ny) then
            print *, "ERROR: restart grid mismatch"
            ok = .false.
            close(unit)
            return
        end if
        read(unit) state%h%data
        read(unit) state%u%data
        read(unit) state%v%data
        close(unit)
        ok = .true.
    end subroutine read_restart
end module swe_io

program test
    use swe_types
    use swe_numerics
    use swe_io
    implicit none
    type(swe_state) :: restart_state, state
    type(grid_t)    :: restart_grid, grid
    logical :: restart_ok

    call restart_grid%init(2, 2, 2.0e6_dp, 2.0e6_dp)
    call restart_state%h%init(2, 2, "h")
    call restart_state%u%init(2, 2, "u")
    call restart_state%v%init(2, 2, "v")
    restart_state%time = 12.0_dp
    restart_state%h%data = 1.0_dp
    restart_state%u%data = 2.0_dp
    restart_state%v%data = 3.0_dp
    call write_restart(restart_state, restart_grid, "swe_restart_tbp_io_negative.bin")

    call grid%init(4, 4, 1.0e6_dp, 1.0e6_dp)
    call state%h%init(4, 4, "h")
    call state%u%init(4, 4, "u")
    call state%v%init(4, 4, "v")
    state%time = 0.0_dp

    call read_restart(state, grid, "swe_restart_tbp_io_negative.bin", restart_ok)

    if (.not. restart_ok) then
        print *, "no restart"
    end if

    print *, "before print"
    call grid%print()
    print *, "after print"
end program test
"#,
    );
    assert_eq!(
        out,
        vec![
            "ERROR: restart grid mismatch",
            "no restart",
            "before print",
            "4",
            "after print"
        ]
    );
}

#[test]
fn tbp_call_after_restart_grid_mismatch() {
    // The real weather_model failure reaches read_restart(), prints
    // "ERROR: restart grid mismatch", then later traps on grid%print().
    // Keep a direct field read before the TBP call so the failure tells us
    // whether the whole object or just the method binding was corrupted.
    let out = run_prints(
        r#"
module swe_types
    implicit none
    integer, parameter :: dp = kind(1.0d0)

    type :: grid_t
        integer  :: nx, ny
        real(dp) :: dx, dy, Lx, Ly
        real(dp), allocatable :: x(:), y(:)
    contains
        procedure :: init  => grid_init
        procedure :: print => grid_print
    end type grid_t

    type :: field2d
        real(dp), allocatable :: data(:,:)
        integer :: nx, ny
        character(len=32) :: name
    contains
        procedure :: init => field_init
    end type field2d

    type :: swe_state
        type(field2d) :: h
        type(field2d) :: u
        type(field2d) :: v
        real(dp) :: time
    end type swe_state

contains
    subroutine grid_init(self, nx, ny, Lx, Ly)
        class(grid_t), intent(inout) :: self
        integer,  intent(in) :: nx, ny
        real(dp), intent(in) :: Lx, Ly
        integer :: i
        self%nx = nx;  self%ny = ny
        self%Lx = Lx;  self%Ly = Ly
        self%dx = Lx / nx
        self%dy = Ly / ny
        allocate(self%x(nx), self%y(ny))
        self%x = [(( i - 0.5_dp) * self%dx, i = 1, nx)]
        self%y = [(( i - 0.5_dp) * self%dy, i = 1, ny)]
    end subroutine grid_init

    subroutine grid_print(self)
        class(grid_t), intent(in) :: self
        print *, self%nx
        print *, self%ny
    end subroutine grid_print

    subroutine field_init(self, nx, ny, name)
        class(field2d), intent(inout) :: self
        integer, intent(in) :: nx, ny
        character(len=*), intent(in) :: name
        self%nx = nx;  self%ny = ny
        self%name = name
        allocate(self%data(nx, ny))
        self%data = 0.0_dp
    end subroutine field_init
end module swe_types

module swe_io
    use swe_types
    implicit none
contains
    subroutine write_restart(state, grid, filename)
        type(swe_state), intent(in) :: state
        type(grid_t),    intent(in) :: grid
        character(len=*), intent(in) :: filename
        integer :: unit
        open(newunit=unit, file=trim(filename), form="unformatted", &
             status="replace", action="write")
        write(unit) state%time
        write(unit) grid%nx, grid%ny
        write(unit) state%h%data
        write(unit) state%u%data
        write(unit) state%v%data
        close(unit)
    end subroutine write_restart

    subroutine read_restart(state, grid, filename, ok)
        type(swe_state), intent(inout) :: state
        type(grid_t),    intent(in)    :: grid
        character(len=*), intent(in)   :: filename
        logical, intent(out) :: ok
        integer :: unit, nx, ny, ios
        open(newunit=unit, file=trim(filename), form="unformatted", &
             status="old", action="read", iostat=ios)
        if (ios /= 0) then
            ok = .false.
            return
        end if
        read(unit) state%time
        read(unit) nx, ny
        if (nx /= grid%nx .or. ny /= grid%ny) then
            print *, "ERROR: restart grid mismatch"
            ok = .false.
            close(unit)
            return
        end if
        read(unit) state%h%data
        read(unit) state%u%data
        read(unit) state%v%data
        close(unit)
        ok = .true.
    end subroutine read_restart
end module swe_io

program test
    use swe_types
    use swe_io
    implicit none
    type(swe_state) :: restart_state, state
    type(grid_t)    :: restart_grid, grid
    logical :: restart_ok

    call restart_grid%init(2, 2, 2.0e6_dp, 2.0e6_dp)
    call restart_state%h%init(2, 2, "h")
    call restart_state%u%init(2, 2, "u")
    call restart_state%v%init(2, 2, "v")
    restart_state%time = 12.0_dp
    restart_state%h%data = 1.0_dp
    restart_state%u%data = 2.0_dp
    restart_state%v%data = 3.0_dp
    call write_restart(restart_state, restart_grid, "swe_restart_mismatch_grid_print.bin")

    call grid%init(4, 4, 1.0e6_dp, 1.0e6_dp)
    call state%h%init(4, 4, "h")
    call state%u%init(4, 4, "u")
    call state%v%init(4, 4, "v")
    state%time = 0.0_dp

    call read_restart(state, grid, "swe_restart_mismatch_grid_print.bin", restart_ok)
    if (.not. restart_ok) then
        print *, "mismatch"
    end if

    print *, grid%nx
    call grid%print()
    print *, "after print"
end program test
"#,
    );
    assert_eq!(
        out,
        vec![
            "ERROR: restart grid mismatch",
            "mismatch",
            "4",
            "4",
            "4",
            "after print"
        ]
    );
}
