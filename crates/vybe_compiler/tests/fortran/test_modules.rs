use super::helpers::compile_ok;

// ═══════════════════════════════════════════════════════════
// Fortran: Modules, use, derived types
// ═══════════════════════════════════════════════════════════

#[test]
fn module_constants() {
    compile_ok("module consts\nreal, parameter :: PI = 3.14159\nend module consts\nprogram t\nuse consts\nprint *, PI\nend program t\n");
}

#[test]
fn module_with_contains() {
    compile_ok("module utils\nimplicit none\ncontains\nfunction sq(x) result(r)\nreal, intent(in) :: x\nreal :: r\nr = x * x\nend function sq\nend module utils\nprogram t\nuse utils\nprint *, sq(5.0)\nend program t\n");
}

#[test]
fn use_only() {
    compile_ok("module mymod\ninteger :: a = 10\ninteger :: b = 20\nend module mymod\nprogram t\nuse mymod, only: a\nprint *, a\nend program t\n");
}

#[test]
fn derived_type_simple() {
    compile_ok("program t\ntype :: Point\nreal :: x\nreal :: y\nend type Point\ntype(Point) :: p\np%x = 3.0\np%y = 4.0\nprint *, p%x\nend program t\n");
}

#[test]
fn derived_type_extends() {
    compile_ok("program t\ntype :: Shape\nreal :: area\nend type Shape\ntype, extends(Shape) :: Circle\nreal :: radius\nend type Circle\ntype(Circle) :: c\nc%radius = 5.0\nprint *, c%radius\nend program t\n");
}

#[test]
fn type_with_procedure() {
    compile_ok("program t\ntype :: Counter\ninteger :: val = 0\ncontains\nprocedure :: inc\nend type Counter\ntype(Counter) :: c\nprint *, c%val\ncontains\nsubroutine inc(self)\nclass(Counter), intent(inout) :: self\nself%val = self%val + 1\nend subroutine inc\nend program t\n");
}

#[test]
fn allocatable_array() {
    compile_ok("program t\ninteger, allocatable :: arr(:)\nallocate(arr(5))\narr(1) = 42\nprint *, arr(1)\ndeallocate(arr)\nend program t\n");
}

#[test]
fn dimension_array() {
    compile_ok("program t\ninteger, dimension(5) :: arr\narr(1) = 10\nprint *, arr(1)\nend program t\n");
}

#[test]
fn interface_block() {
    compile_ok("program t\ninterface\nfunction add(a, b) result(r)\ninteger, intent(in) :: a, b\ninteger :: r\nend function add\nend interface\nprint *, \"ok\"\nend program t\n");
}
