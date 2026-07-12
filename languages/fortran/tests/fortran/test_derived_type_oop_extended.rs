//! Extended derived-type OOP: EXTENDS component shadowing, SELECT TYPE guards,
//! allocatable-component lifecycle, pointer components, and TBP overrides.
//! Distinct from `test_derived_types_advanced.rs` (constructors, weather-model TBP,
//! basic SELECT TYPE, allocatable type arrays).

use super::helpers::compile_ok;

fortran_cases! {
    // ── EXTENDS with component shadowing ─────────────────────────────

    extends_shadow_child_integer_wins => {
        "program t\ntype :: Base\ninteger :: val = 1\nend type Base\ntype, extends(Base) :: Derived\ninteger :: val = 99\nend type Derived\ntype(Derived) :: d\nprint *, d%val\nend program t\n",
        ["99"]
    };

    extends_shadow_child_real_overrides_parent => {
        "program t\ntype :: Base\nreal :: metric = 1.0\nend type Base\ntype, extends(Base) :: Derived\nreal :: metric = 3.5\nend type Derived\ntype(Derived) :: d\nprint *, int(d%metric)\nend program t\n",
        ["3"]
    };

    // ── Allocatable component lifecycle ──────────────────────────────

    alloc_comp_deallocate_clears_component => {
        "program t\ntype :: Bucket\ninteger, allocatable :: slots(:)\nend type Bucket\ntype(Bucket) :: b\nallocate(b%slots(2))\nb%slots = [4, 6]\nprint *, b%slots(1)\ndeallocate(b%slots)\nprint *, allocated(b%slots)\nend program t\n",
        ["4", "false"]
    };

    alloc_comp_reallocate_after_deallocate => {
        "program t\ntype :: Buffer\ninteger, allocatable :: data(:)\nend type Buffer\ntype(Buffer) :: buf\nallocate(buf%data(2))\nbuf%data = [1, 2]\ndeallocate(buf%data)\nallocate(buf%data(3))\nbuf%data = [3, 4, 5]\nprint *, buf%data(3)\nend program t\n",
        ["5"]
    };

    alloc_comp_allocated_status_on_member => {
        "program t\ntype :: Holder\nreal, allocatable :: vals(:)\nend type Holder\ntype(Holder) :: h\nprint *, allocated(h%vals)\nallocate(h%vals(1))\nh%vals(1) = 2.5\nprint *, int(h%vals(1))\nprint *, allocated(h%vals)\nend program t\n",
        ["false", "2", "true"]
    };

    // ── Pointer components ───────────────────────────────────────────

    dertype_pointer_targets_scalar_member => {
        "program t\ntype :: Node\ninteger :: key = 0\ninteger, pointer :: link => null()\nend type Node\ntype(Node), target :: a, b\na%key = 11\nb%key = 22\na%link => b\nprint *, a%link%key\nend program t\n",
        ["22"]
    };

    dertype_pointer_reassign_to_new_target => {
        "program t\ntype :: Pair\ninteger :: left = 0\ninteger, pointer :: right => null()\nend type Pair\ntype(Pair) :: p\ninteger, target :: x = 3, y = 8\np%left = 1\np%right => x\nprint *, p%right\np%right => y\nprint *, p%right\nend program t\n",
        ["3", "8"]
    };

    // ── Type-bound procedure overrides (static dispatch) ─────────────

    tbp_child_overrides_parent_tag_function => {
        "program t\ntype :: Base\ncontains\nprocedure :: tag\nend type Base\ntype, extends(Base) :: Child\ncontains\nprocedure :: tag => child_tag\nend type Child\ntype(Child) :: c\nprint *, c%tag()\ncontains\ninteger function tag(self) result(v)\nclass(Base), intent(in) :: self\nv = 1\nend function tag\ninteger function child_tag(self) result(v)\nclass(Child), intent(in) :: self\nv = 2\nend function child_tag\nend program t\n",
        ["2"]
    };
}

// ── TYPE IS / CLASS IS guards (compile-only) ────────────────────────

#[test]
fn compile_select_type_is_child_branch() {
    compile_ok(
        r#"
program t
    type :: Base
        integer :: id = 0
    end type Base
    type, extends(Base) :: Child
        integer :: extra = 42
    end type Child
    class(Base), allocatable :: obj
    allocate(Child :: obj)
    select type(obj)
    type is (Child)
        print *, obj%extra
    class default
        print *, obj%id
    end select
end program t
"#,
    );
}

#[test]
fn compile_select_class_is_extended_type() {
    compile_ok(
        r#"
program t
    type :: A
        integer :: x = 1
    end type A
    type, extends(A) :: B
        integer :: y = 7
    end type B
    class(A), allocatable :: obj
    allocate(B :: obj)
    select type(obj)
    class is (B)
        print *, obj%y
    type is (A)
        print *, obj%x
    end select
end program t
"#,
    );
}

#[test]
fn compile_select_type_is_exact_base_type() {
    compile_ok(
        r#"
program t
    type :: Base
        integer :: id = 5
    end type Base
    type, extends(Base) :: Child
        integer :: extra = 9
    end type Child
    class(Base), allocatable :: obj
    allocate(Base :: obj)
    select type(obj)
    type is (Base)
        print *, obj%id
    class is (Child)
        print *, obj%extra
    class default
        print *, 0
    end select
end program t
"#,
    );
}

#[test]
fn compile_select_type_class_default_guard() {
    compile_ok(
        r#"
program t
    type :: Root
        integer :: tag = 1
    end type Root
    type, extends(Root) :: Leaf
        integer :: payload = 4
    end type Leaf
    class(Root), allocatable :: node
    allocate(Root :: node)
    select type(node)
    class is (Leaf)
        print *, node%payload
    type is (Root)
        print *, node%tag
    class default
        print *, 0
    end select
end program t
"#,
    );
}

// ── Type-bound procedure overrides (polymorphic dispatch) ───────────

#[test]
fn compile_tbp_polymorphic_child_override_dispatch() {
    compile_ok(
        r#"
program t
    type :: Animal
    contains
        procedure :: legs
    end type Animal
    type, extends(Animal) :: Spider
    contains
        procedure :: legs => spider_legs
    end type Spider
    class(Animal), allocatable :: a
    allocate(Spider :: a)
    print *, a%legs()
contains
    integer function legs(self) result(n)
        class(Animal), intent(in) :: self
        n = 4
    end function legs
    integer function spider_legs(self) result(n)
        class(Spider), intent(in) :: self
        n = 8
    end function spider_legs
end program t
"#,
    );
}

// ── FINAL procedures (compile-only) ─────────────────────────────────

#[test]
fn compile_final_on_extended_child_type() {
    compile_ok(
        r#"
program t
    type :: Base
        integer :: id = 0
    contains
        final :: base_done
    end type Base
    type, extends(Base) :: Child
    contains
        final :: child_done
    end type Child
    type(Child) :: c
    c%id = 7
    print *, c%id
contains
    subroutine base_done(self)
        type(Base), intent(inout) :: self
        self%id = 0
    end subroutine base_done
    subroutine child_done(self)
        type(Child), intent(inout) :: self
        self%id = -1
    end subroutine child_done
end program t
"#,
    );
}

#[test]
fn compile_final_multiple_procedures_in_type() {
    compile_ok(
        r#"
program t
    type :: Resource
        integer :: handle = 0
        logical :: open = .false.
    contains
        final :: close_handle
        final :: mark_closed
    end type Resource
    type(Resource) :: r
    r%handle = 42
    r%open = .true.
    print *, r%handle
contains
    subroutine close_handle(self)
        type(Resource), intent(inout) :: self
        self%handle = 0
    end subroutine close_handle
    subroutine mark_closed(self)
        type(Resource), intent(inout) :: self
        self%open = .false.
    end subroutine mark_closed
end program t
"#,
    );
}

// ── Pointer component nullify (compile-only) ────────────────────────

#[test]
fn compile_dertype_pointer_nullify_component() {
    compile_ok(
        r#"
program t
    type :: Link
        integer :: value = 0
        type(Link), pointer :: next => null()
    end type Link
    type(Link), target :: head, tail
    head%value = 1
    tail%value = 2
    head%next => tail
    nullify(head%next)
    print *, head%value
end program t
"#,
    );
}
