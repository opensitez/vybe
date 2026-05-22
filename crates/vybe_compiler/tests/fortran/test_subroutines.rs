use super::helpers::{compile_ok, run_prints};
use vybe_compiler::ast::StmtKind;

// ═══════════════════════════════════════════════════════════
// Fortran: Subroutines and functions (contains block)
// ═══════════════════════════════════════════════════════════

#[test]
fn subroutine_empty() {
    compile_ok("program t\ncall greet()\ncontains\nsubroutine greet()\nprint *, \"hi\"\nend subroutine greet\nend program t\n");
}

#[test]
fn subroutine_with_arg() {
    compile_ok("program t\ncall say(\"hello\")\ncontains\nsubroutine say(msg)\ncharacter(len=*), intent(in) :: msg\nprint *, msg\nend subroutine say\nend program t\n");
}

#[test]
fn function_returns_value() {
    compile_ok("program t\nprint *, double(5)\ncontains\nfunction double(x) result(res)\ninteger, intent(in) :: x\ninteger :: res\nres = x * 2\nend function double\nend program t\n");
}

#[test]
fn function_with_type_prefix() {
    compile_ok("program t\nprint *, add(3, 4)\ncontains\ninteger function add(a, b)\ninteger, intent(in) :: a, b\nadd = a + b\nend function add\nend program t\n");
}

#[test]
fn multiple_contains() {
    compile_ok("program t\ncontains\nsubroutine a()\nprint *, \"a\"\nend subroutine a\nsubroutine b()\nprint *, \"b\"\nend subroutine b\nend program t\n");
}

#[test]
fn recursive_factorial() {
    compile_ok("program t\nprint *, fact(5)\ncontains\nrecursive function fact(n) result(r)\ninteger, intent(in) :: n\ninteger :: r\nif (n <= 1) then\nr = 1\nelse\nr = n * fact(n - 1)\nend if\nend function fact\nend program t\n");
}

#[test]
fn procedure_dummy_param_is_stamped_callable() {
    let src = "module m\n  implicit none\n  abstract interface\n    function rhs_func(t) result(v)\n      real, intent(in) :: t\n      real :: v\n    end function rhs_func\n  end interface\ncontains\n  subroutine step(rhs)\n    procedure(rhs_func) :: rhs\n    print *, rhs(2.0)\n  end subroutine step\nend module m\n";
    let module = vybe_compiler::languages::fortran::parse(src).expect("parse failed");
    let step = module
        .body
        .iter()
        .find_map(|statement| match &statement.kind {
            StmtKind::ModuleDecl { members, .. } => members.iter().find_map(|member| match member {
                vybe_compiler::ast::ClassMember::Method(stmt) => match &stmt.kind {
                    StmtKind::FunctionDecl { name, params, .. } if name.eq_ignore_ascii_case("step") => Some(params),
                    _ => None,
                },
                _ => None,
            }),
            _ => None,
        })
        .expect("missing step params");

    let rhs = step
        .iter()
        .find(|param| param.name.eq_ignore_ascii_case("rhs"))
        .expect("missing rhs param");

    assert_eq!(rhs.type_hint.as_deref(), Some("procedure(rhs_func)"));
}

#[test]
fn procedure_dummy_call_uses_abstract_interface_signature() {
    let out = run_prints(
        "module m\n  implicit none\n  abstract interface\n    integer function rhs_func(t) result(v)\n      integer, intent(in) :: t\n      integer :: v\n    end function rhs_func\n  end interface\ncontains\n  subroutine step(rhs)\n    procedure(rhs_func) :: rhs\n    print *, rhs(2)\n  end subroutine step\nend module m\n\nprogram test\n  use m\n  call step(rhs1)\ncontains\n  integer function rhs1(t) result(v)\n    integer, intent(in) :: t\n    integer :: v\n    v = t * 2\n  end function rhs1\nend program test\n",
    );

    assert_eq!(out, ["4"]);
}

#[test]
fn interface_procedure_member_aliases_signature() {
    let out = run_prints(
        "module m\n  implicit none\n  abstract interface\n    integer function unary(x) result(v)\n      integer, intent(in) :: x\n      integer :: v\n    end function unary\n    procedure(unary) :: op\n  end interface\ncontains\n  subroutine step(rhs)\n    procedure(op) :: rhs\n    print *, rhs(3)\n  end subroutine step\nend module m\n\nprogram test\n  use m\n  call step(double_it)\ncontains\n  integer function double_it(x) result(v)\n    integer, intent(in) :: x\n    integer :: v\n    v = x * 2\n  end function double_it\nend program test\n",
    );

    assert_eq!(out, ["6"]);
}

#[test]
fn procedure_dummy_array_result_preserves_values() {
    let out = run_prints(
        "module m\n  implicit none\n  abstract interface\n    function rhs_func(t, y, n) result(dydt)\n      integer, intent(in) :: n\n      real, intent(in) :: t\n      real, intent(in) :: y(n)\n      real :: dydt(n)\n    end function rhs_func\n  end interface\ncontains\n  subroutine step(rhs)\n    procedure(rhs_func) :: rhs\n    real :: y(3), dydt(3)\n    y = [1.0, 0.0, 0.0]\n    dydt = rhs(0.0, y, 3)\n    print *, dydt(1)\n    print *, dydt(2)\n    print *, dydt(3)\n  end subroutine step\nend module m\n\nprogram test\n  use m\n  call step(lorenz_rhs)\ncontains\n  function lorenz_rhs(t, y, n) result(dydt)\n    integer, intent(in) :: n\n    real, intent(in) :: t\n    real, intent(in) :: y(n)\n    real :: dydt(n)\n    dydt(1) = 10.0 * (y(2) - y(1))\n    dydt(2) = y(1) * (28.0 - y(3)) - y(2)\n    dydt(3) = y(1) * y(2)\n  end function lorenz_rhs\nend program test\n",
    );

    assert_eq!(out, ["-10", "28", "0"]);
}

#[test]
fn procedure_dummy_array_result_advances_state_time() {
    let out = run_prints(
        "module m\n  implicit none\n  type :: ode_state\n    real :: t\n    real :: y(3)\n  end type ode_state\n  abstract interface\n    function rhs_func(t, y, n) result(dydt)\n      integer, intent(in) :: n\n      real, intent(in) :: t\n      real, intent(in) :: y(n)\n      real :: dydt(n)\n    end function rhs_func\n  end interface\ncontains\n  subroutine rk4_step(state, h, rhs)\n    type(ode_state), intent(inout) :: state\n    real, intent(in) :: h\n    procedure(rhs_func) :: rhs\n    real :: k1(3)\n    k1 = rhs(state%t, state%y, 3)\n    state%y = state%y + h * k1\n    state%t = state%t + h\n  end subroutine rk4_step\nend module m\n\nprogram test\n  use m\n  type(ode_state) :: state\n  state%t = 0.0\n  state%y = [1.0, 0.0, 0.0]\n  call rk4_step(state, 0.5, lorenz_rhs)\n  print *, state%t\n  print *, state%y(1)\n  print *, state%y(2)\n  print *, state%y(3)\ncontains\n  function lorenz_rhs(t, y, n) result(dydt)\n    integer, intent(in) :: n\n    real, intent(in) :: t\n    real, intent(in) :: y(n)\n    real :: dydt(n)\n    dydt(1) = 10.0 * (y(2) - y(1))\n    dydt(2) = y(1) * (28.0 - y(3)) - y(2)\n    dydt(3) = y(1) * y(2)\n  end function lorenz_rhs\nend program test\n",
    );

    assert_eq!(out, ["0.5", "-4", "14", "0"]);
}

#[test]
fn array_field_assignment_is_lowered_to_element_loop() {
    let module = vybe_compiler::languages::fortran::parse(
        "module m\n  implicit none\n  type :: ode_state\n    real :: t\n    real :: y(3)\n  end type ode_state\n  abstract interface\n    function rhs_func(t, y, n) result(dydt)\n      integer, intent(in) :: n\n      real, intent(in) :: t\n      real, intent(in) :: y(n)\n      real :: dydt(n)\n    end function rhs_func\n  end interface\ncontains\n  subroutine rk4_step(state, h, rhs)\n    type(ode_state), intent(inout) :: state\n    real, intent(in) :: h\n    procedure(rhs_func) :: rhs\n    real :: k1(3)\n    k1 = rhs(state%t, state%y, 3)\n    state%y = state%y + h * k1\n    state%t = state%t + h\n  end subroutine rk4_step\nend module m\n",
    )
    .expect("parse failed");

    let lowered = module
        .body
        .iter()
        .find_map(|statement| match &statement.kind {
            StmtKind::ModuleDecl { members, .. } => members.iter().find_map(|member| match member {
                vybe_compiler::ast::ClassMember::Method(stmt) => match &stmt.kind {
                    StmtKind::FunctionDecl { name, body, .. } if name.eq_ignore_ascii_case("rk4_step") => {
                        Some(body.iter().any(|stmt| {
                            matches!(&stmt.kind, StmtKind::Block(stmts) if stmts.iter().any(|inner| matches!(inner.kind, StmtKind::For { .. })))
                        }))
                    }
                    _ => None,
                },
                _ => None,
            }),
            _ => None,
        })
        .expect("missing rk4_step body");

    assert!(lowered, "expected state%y array assignment to lower into a loop");
}

#[test]
fn derived_type_fixed_array_field_supports_index_read_and_write() {
    let out = run_prints(
        "program test\n  type :: ode_state\n    real :: y(3)\n  end type ode_state\n  type(ode_state) :: state\n  state%y = [1.0, 0.0, 0.0]\n  print *, state%y(1)\n  state%y(2) = 4.0\n  print *, state%y(2)\nend program test\n",
    );

    assert_eq!(out, ["1", "4"]);
}

#[test]
fn derived_type_array_literal_assignment_is_lowered_to_element_loop() {
    let module = vybe_compiler::languages::fortran::parse(
        "program test\n  type :: ode_state\n    real :: y(3)\n  end type ode_state\n  type(ode_state) :: state\n  state%y = [1.0, 0.0, 0.0]\nend program test\n",
    )
    .expect("parse failed");

    let lowered = module.body.iter().any(|statement| {
        matches!(&statement.kind, StmtKind::Block(stmts) if stmts.iter().any(|inner| matches!(inner.kind, StmtKind::For { .. })))
    });

    assert!(lowered, "expected derived-type array literal assignment to lower into a loop");
}

#[test]
fn derived_type_allocatable_array_field_supports_rk4_array_math() {
    let out = run_prints(
        "module m\n  implicit none\n  type :: ode_state\n    real :: t\n    real, allocatable :: y(:)\n    integer :: neq\n  contains\n    procedure :: init => state_init\n  end type ode_state\n  abstract interface\n    function rhs_func(t, y, n) result(dydt)\n      integer, intent(in) :: n\n      real, intent(in) :: t\n      real, intent(in) :: y(n)\n      real :: dydt(n)\n    end function rhs_func\n  end interface\ncontains\n  subroutine state_init(self, neq, t0)\n    class(ode_state), intent(inout) :: self\n    integer, intent(in) :: neq\n    real, intent(in) :: t0\n    self%neq = neq\n    self%t = t0\n    allocate(self%y(neq))\n    self%y = 0.0\n  end subroutine\n\n  subroutine rk4_step(state, h, rhs)\n    type(ode_state), intent(inout) :: state\n    real, intent(in) :: h\n    procedure(rhs_func) :: rhs\n    real :: k1(state%neq), k2(state%neq), k3(state%neq), k4(state%neq), h2\n    h2 = h * 0.5\n    k1 = rhs(state%t, state%y, state%neq)\n    k2 = rhs(state%t + h2, state%y + h2 * k1, state%neq)\n    k3 = rhs(state%t + h2, state%y + h2 * k2, state%neq)\n    k4 = rhs(state%t + h, state%y + h * k3, state%neq)\n    state%y = state%y + (h / 6.0) * (k1 + 2.0 * k2 + 2.0 * k3 + k4)\n    state%t = state%t + h\n  end subroutine\nend module m\n\nprogram test\n  use m\n  type(ode_state) :: state\n  call state%init(3, 0.0)\n  state%y = [1.0, 0.0, 0.0]\n  call rk4_step(state, 0.01, lorenz_rhs)\n  print *, state%t\n  print *, state%y(1)\n  print *, state%y(2)\n  print *, state%y(3)\ncontains\n  function lorenz_rhs(t, y, n) result(dydt)\n    integer, intent(in) :: n\n    real, intent(in) :: t\n    real, intent(in) :: y(n)\n    real :: dydt(n)\n    dydt(1) = 10.0 * (y(2) - y(1))\n    dydt(2) = y(1) * (28.0 - y(3)) - y(2)\n    dydt(3) = y(1) * y(2)\n  end function lorenz_rhs\nend program test\n",
    );

    assert_eq!(out[0], "0.01");
    assert!(out[1..].iter().all(|line| !line.contains("NaN") && !line.contains(',')), "expected scalar allocatable field elements after RK4 step, got {out:?}");
}

#[test]
fn nested_allocatable_member_index_assignment_writes_back_receiver() {
    let out = run_prints(
        "program test\n  type :: field2d\n    real, allocatable :: data(:,:)\n  contains\n    procedure :: init => field_init\n  end type field2d\n  type :: state_t\n    type(field2d) :: h\n  end type state_t\n  type(state_t) :: state\n  call state%h%init(3, 4)\n  state%h%data(2, 3) = 42.0\n  state%h%data(1, 1) = -7.5\n  print *, state%h%data(2, 3)\n  print *, state%h%data(1, 1)\ncontains\n  subroutine field_init(self, nx, ny)\n    class(field2d), intent(inout) :: self\n    integer, intent(in) :: nx, ny\n    allocate(self%data(nx, ny))\n    self%data = 0.0\n  end subroutine field_init\nend program test\n",
    );

    assert_eq!(out, ["42", "-7.5"]);
}

#[test]
fn multidimensional_scalar_broadcast_preserves_row_slices() {
    let out = run_prints(
        "program test\n  integer :: i\n  real :: trajectory(4, 5)\n  trajectory = 0.0\n  trajectory(2, :) = [3.0, -2.0, 7.0, 1.5, 4.0]\n  i = 1\n  print *, trajectory(2)\n  print *, trajectory(2, 1)\n  print *, minval(trajectory(i + 1, 1:5))\n  print *, maxval(trajectory(i + 1, 1:5))\nend program test\n",
    );

    assert_eq!(out, ["3,-2,7,1.5,4", "3", "-2", "7"]);
}

#[test]
fn multidimensional_array_math_preserves_nested_shape() {
    let out = run_prints(
        "program test\n  real, allocatable :: a(:,:), b(:,:), c(:,:)\n  allocate(a(2,2), b(2,2), c(2,2))\n  a = 1.0\n  b = 2.0\n  c = a + b\n  print *, c(1,1)\n  print *, c(2,2)\n  print *, sum(c)\nend program test\n",
    );

    assert_eq!(out, ["3", "3", "12"]);
}

#[test]
fn multidimensional_member_array_reductions_cover_all_elements() {
    let out = run_prints(
        "program test\n  type :: field2d\n    real, allocatable :: data(:,:)\n  contains\n    procedure :: init => field_init\n  end type\n  type(field2d) :: h, u\n  call h%init(2,2)\n  call u%init(2,2)\n  h%data = 3.0\n  u%data = 4.0\n  print *, size(h%data)\n  print *, maxval(h%data * u%data)\n  print *, sum(h%data * u%data)\ncontains\n  subroutine field_init(self, nx, ny)\n    class(field2d), intent(inout) :: self\n    integer, intent(in) :: nx, ny\n    allocate(self%data(nx, ny))\n    self%data = 0.0\n  end subroutine field_init\nend program test\n",
    );

    assert_eq!(out, ["4", "12", "48"]);
}

#[test]
fn explicit_shape_parameter_whole_array_unary_assignment_preserves_shape() {
    let out = run_prints(
        "program test\n  real :: a(2,2), b(2,2)\n  call foo(a, b)\n  print *, sum(a)\n  print *, sum(b)\ncontains\n  subroutine foo(a, b)\n    real, intent(out) :: a(2,2)\n    real, intent(out) :: b(2,2)\n    a = 4.0\n    b = -a\n  end subroutine foo\nend program test\n",
    );

    assert_eq!(out, ["16", "-16"]);
}

#[test]
fn explicit_shape_vector_parameter_unary_assignment_preserves_values() {
    let out = run_prints(
        "program test\n  real :: a(4), b(4)\n  call foo(a, b)\n  print *, sum(a)\n  print *, sum(b)\ncontains\n  subroutine foo(a, b)\n    real, intent(out) :: a(4)\n    real, intent(out) :: b(4)\n    a = 4.0\n    b = -a\n  end subroutine foo\nend program test\n",
    );

    assert_eq!(out, ["16", "-16"]);
}

#[test]
fn allocatable_member_array_copy_preserves_values() {
    let out = run_prints(
        "program test\n  type :: box\n    integer :: n\n    real, allocatable :: y(:)\n  contains\n    procedure :: init => box_init\n    procedure :: copy => box_copy\n  end type\n  type(box) :: a, b\n  call a%init(3)\n  a%y = [1.0, 2.0, 3.0]\n  call b%copy(a)\n  print *, sum(b%y)\ncontains\n  subroutine box_init(self, n)\n    class(box), intent(inout) :: self\n    integer, intent(in) :: n\n    self%n = n\n    allocate(self%y(n))\n    self%y = 0.0\n  end subroutine box_init\n  subroutine box_copy(self, other)\n    class(box), intent(inout) :: self\n    type(box), intent(in) :: other\n    self%n = other%n\n    if (allocated(self%y)) deallocate(self%y)\n    allocate(self%y(other%n))\n    self%y = other%y\n  end subroutine box_copy\nend program test\n",
    );

    assert_eq!(out, ["6"]);
}
