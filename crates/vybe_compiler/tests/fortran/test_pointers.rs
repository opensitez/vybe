use super::helpers::compile_ok;

// ── Pointer declarations ──────────────────────────────────────

#[test] fn pointer_integer() { compile_ok("program t\n  integer, pointer :: p => null()\n  print *, associated(p)\nend program t\n"); }
#[test] fn pointer_real() { compile_ok("program t\n  real, pointer :: p => null()\n  print *, associated(p)\nend program t\n"); }
#[test] fn pointer_char() { compile_ok("program t\n  character(len=10), pointer :: p => null()\n  print *, associated(p)\nend program t\n"); }
#[test] fn pointer_logical() { compile_ok("program t\n  logical, pointer :: p => null()\n  print *, associated(p)\nend program t\n"); }

// ── TARGET attribute ──────────────────────────────────────────

#[test] fn target_integer() { compile_ok("program t\n  integer, target :: x = 42\n  integer, pointer :: p\n  p => x\n  print *, p\nend program t\n"); }
#[test] fn target_real() { compile_ok("program t\n  real, target :: x = 3.14\n  real, pointer :: p\n  p => x\n  print *, p\nend program t\n"); }
#[test] fn target_array() { compile_ok("program t\n  integer, target :: a(5) = [1,2,3,4,5]\n  integer, pointer :: p(:)\n  p => a\n  print *, p(1)\nend program t\n"); }

// ── Pointer association ───────────────────────────────────────

#[test] fn pointer_associate() {
    compile_ok(r#"
program test
    integer, target :: x = 10
    integer, pointer :: p
    p => x
    print *, p
end program test
"#);
}

#[test] fn pointer_deref_modify() {
    compile_ok(r#"
program test
    integer, target :: x = 10
    integer, pointer :: p
    p => x
    p = 99
    print *, x
end program test
"#);
}

#[test] fn pointer_reassociate() {
    compile_ok(r#"
program test
    integer, target :: a = 1, b = 2
    integer, pointer :: p
    p => a
    p => b
    print *, p
end program test
"#);
}

// ── ASSOCIATED intrinsic ──────────────────────────────────────

#[test] fn associated_null() {
    compile_ok(r#"
program test
    integer, pointer :: p => null()
    print *, associated(p)
end program test
"#);
}

#[test] fn associated_after_target() {
    compile_ok(r#"
program test
    integer, target :: x = 5
    integer, pointer :: p
    p => x
    print *, associated(p)
end program test
"#);
}

#[test] fn associated_with_target() {
    compile_ok(r#"
program test
    integer, target :: x = 1, y = 2
    integer, pointer :: p
    p => x
    print *, associated(p, x)
    print *, associated(p, y)
end program test
"#);
}

// ── NULLIFY ───────────────────────────────────────────────────

#[test] fn nullify_basic() {
    compile_ok(r#"
program test
    integer, pointer :: p => null()
    integer, target :: x = 5
    p => x
    nullify(p)
    print *, associated(p)
end program test
"#);
}

#[test] fn nullify_multiple() {
    compile_ok(r#"
program test
    integer, pointer :: p => null(), q => null()
    integer, target :: x = 1, y = 2
    p => x
    q => y
    nullify(p, q)
    print *, associated(p)
end program test
"#);
}

// ── Pointer arrays ────────────────────────────────────────────

#[test] fn pointer_array_1d() {
    compile_ok(r#"
program test
    integer, target :: a(5) = [10, 20, 30, 40, 50]
    integer, pointer :: p(:)
    p => a
    print *, p(3)
end program test
"#);
}

#[test] fn pointer_array_slice() {
    compile_ok(r#"
program test
    integer, target :: a(6) = [1, 2, 3, 4, 5, 6]
    integer, pointer :: p(:)
    p => a(2:5)
    print *, p(1)
end program test
"#);
}

#[test] fn pointer_array_2d() {
    compile_ok(r#"
program test
    integer, target :: m(3,3)
    integer, pointer :: p(:,:)
    m = 0
    m(2,2) = 42
    p => m
    print *, p(2,2)
end program test
"#);
}

// ── Pointer in derived type ───────────────────────────────────

#[test] fn pointer_field() {
    compile_ok(r#"
program test
    type :: Node
        integer :: value
        type(Node), pointer :: next => null()
    end type Node
    type(Node) :: n
    n%value = 1
    print *, n%value
end program test
"#);
}

#[test] fn linked_list_two_nodes() {
    compile_ok(r#"
program test
    type :: Node
        integer :: value
        type(Node), pointer :: next => null()
    end type Node
    type(Node), target :: n1, n2
    n1%value = 1
    n2%value = 2
    n1%next => n2
    print *, n1%next%value
end program test
"#);
}

// ── Allocatable with pointer-like semantics ───────────────────

#[test] fn allocatable_scalar() {
    compile_ok(r#"
program test
    integer, allocatable :: x
    allocate(x)
    x = 42
    print *, x
    deallocate(x)
end program test
"#);
}

#[test] fn allocatable_real_scalar() {
    compile_ok(r#"
program test
    real, allocatable :: r
    allocate(r)
    r = 3.14
    print *, r
    deallocate(r)
end program test
"#);
}

#[test] fn allocatable_in_type() {
    compile_ok(r#"
program test
    type :: DynList
        integer, allocatable :: items(:)
        integer :: count = 0
    end type DynList
    type(DynList) :: list
    allocate(list%items(10))
    list%items(1) = 100
    list%count = 1
    print *, list%items(1)
    deallocate(list%items)
end program test
"#);
}

// ── SOURCE= and MOLD= in ALLOCATE (Fortran 2003) ─────────────

#[test] fn allocate_source() {
    compile_ok(r#"
program test
    integer, allocatable :: a(:), b(:)
    a = [1, 2, 3, 4, 5]
    allocate(b, source=a)
    print *, b(1)
    deallocate(a, b)
end program test
"#);
}

#[test] fn allocate_mold() {
    compile_ok(r#"
program test
    integer, allocatable :: a(:), b(:)
    allocate(a(5))
    allocate(b, mold=a)
    print *, size(b)
    deallocate(a, b)
end program test
"#);
}

// ── Procedure pointers (Fortran 2003) ────────────────────────

#[test] fn procedure_pointer_basic() {
    compile_ok(r#"
program test
    procedure(int_fn), pointer :: fp => null()
    fp => double_it
    print *, fp(5)
contains
    function double_it(x) result(r)
        integer, intent(in) :: x
        integer :: r
        r = x * 2
    end function double_it
    function int_fn(x) result(r)
        integer, intent(in) :: x
        integer :: r
    end function int_fn
end program test
"#);
}

#[test] fn procedure_pointer_swap() {
    compile_ok(r#"
program test
    abstract interface
        function unary(x) result(r)
            integer, intent(in) :: x
            integer :: r
        end function unary
    end interface
    procedure(unary), pointer :: fp
    fp => double_it
    print *, fp(3)
    fp => triple_it
    print *, fp(3)
contains
    function double_it(x) result(r)
        integer, intent(in) :: x
        integer :: r
        r = x * 2
    end function double_it
    function triple_it(x) result(r)
        integer, intent(in) :: x
        integer :: r
        r = x * 3
    end function triple_it
end program test
"#);
}
