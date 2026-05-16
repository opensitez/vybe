use super::helpers::compile_ok;

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
