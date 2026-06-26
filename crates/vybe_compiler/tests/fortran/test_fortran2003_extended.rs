//! Extended Fortran 2003 coverage: type-bound procedures, generic bindings,
//! deferred/abstract types, allocatable components, flush, and stream I/O.
//! Distinct from `test_fortran2003.rs` (core F2003) and `test_io_advanced.rs`.

use super::helpers::compile_ok;

fortran_cases! {
    // ── Type-bound procedures ────────────────────────────────────────

    tbp_accumulator_add_twice => {
        "program t\ntype :: Acc\ninteger :: total = 0\ncontains\nprocedure :: bump\nend type Acc\ntype(Acc) :: a\ncall a%bump(5)\ncall a%bump(3)\nprint *, a%total\ncontains\nsubroutine bump(self, n)\nclass(Acc), intent(inout) :: self\ninteger, intent(in) :: n\nself%total = self%total + n\nend subroutine bump\nend program t\n",
        ["8"]
    };

    tbp_bound_function_hypotenuse => {
        "program t\ntype :: Legs\nreal :: a, b\ncontains\nprocedure :: hyp\nend type Legs\ntype(Legs) :: tri\ntri%a = 3.0\ntri%b = 4.0\nprint *, int(tri%hyp())\ncontains\nfunction hyp(self) result(h)\nclass(Legs), intent(in) :: self\nreal :: h\nh = sqrt(self%a**2 + self%b**2)\nend function hyp\nend program t\n",
        ["5"]
    };

    tbp_pass_binding_alias_print => {
        "program t\ntype :: Label\ncharacter(len=8) :: text = 'vybe'\ncontains\nprocedure :: show => emit_label\nend type Label\ntype(Label) :: item\ncall item%show()\ncontains\nsubroutine emit_label(self)\nclass(Label), intent(in) :: self\nprint *, trim(self%text)\nend subroutine emit_label\nend program t\n",
        ["vybe"]
    };

    tbp_module_reset_counter => {
        "module tallies\nimplicit none\ntype :: Tally\ninteger :: n = 0\ncontains\nprocedure :: reset\nend type Tally\ncontains\nsubroutine reset(self)\nclass(Tally), intent(inout) :: self\nself%n = 0\nend subroutine reset\nend module tallies\nprogram t\nuse tallies\ntype(Tally) :: t\nt%n = 9\ncall t%reset()\nprint *, t%n\nend program t\n",
        ["0"]
    };

    tbp_nested_point_distance_origin => {
        "program t\ntype :: Coord\nreal :: x, y\ncontains\nprocedure :: len\nend type Coord\ntype :: Segment\ntype(Coord) :: start, finish\ncontains\nprocedure :: span\nend type Segment\ntype(Segment) :: s\ns%start%x = 0.0\ns%start%y = 0.0\ns%finish%x = 3.0\ns%finish%y = 4.0\nprint *, int(s%span())\ncontains\nfunction len(self) result(d)\nclass(Coord), intent(in) :: self\nreal :: d\nd = sqrt(self%x**2 + self%y**2)\nend function len\nfunction span(self) result(d)\nclass(Segment), intent(in) :: self\nreal :: d\nd = sqrt((self%finish%x - self%start%x)**2 + (self%finish%y - self%start%y)**2)\nend function span\nend program t\n",
        ["5"]
    };

    tbp_string_upper_via_binding => {
        "program t\ntype :: Word\ncharacter(len=12) :: text = 'fortran'\ncontains\nprocedure :: upper\nend type Word\ntype(Word) :: w\nprint *, trim(w%upper())\ncontains\nfunction upper(self) result(out)\nclass(Word), intent(in) :: self\ncharacter(len=12) :: out\nout = self%text\nend function upper\nend program t\n",
        ["fortran"]
    };

    tbp_child_extends_parent_binding => {
        "program t\ntype :: Base\ninteger :: n = 1\ncontains\nprocedure :: twice\nend type Base\ntype, extends(Base) :: Child\ninteger :: extra = 0\nend type Child\ntype(Child) :: c\nprint *, c%twice()\ncontains\nfunction twice(self) result(v)\nclass(Base), intent(in) :: self\ninteger :: v\nv = self%n * 2\nend function twice\nend program t\n",
        ["2"]
    };

    tbp_nopass_scale_helper => {
        "program t\ntype :: Scale\ncontains\nprocedure(scale_by), nopass :: apply\nend type Scale\ntype(Scale) :: s\nprint *, s%apply(6, 2)\ncontains\ninteger function scale_by(n, k) result(r)\ninteger, intent(in) :: n, k\nr = n * k\nend function scale_by\nend program t\n",
        ["12"]
    };

    // ── Generic type-bound bindings ──────────────────────────────────

    generic_int_show_binding => {
        "program t\ntype :: Printer\ncontains\nprocedure :: show_i\nprocedure :: show_r\ngeneric :: show => show_i, show_r\nend type Printer\ntype(Printer) :: p\ncall p%show(7)\ncall p%show(2.5)\ncontains\nsubroutine show_i(self, v)\nclass(Printer), intent(in) :: self\ninteger, intent(in) :: v\nprint *, v\nend subroutine show_i\nsubroutine show_r(self, v)\nclass(Printer), intent(in) :: self\nreal, intent(in) :: v\nprint *, int(v)\nend subroutine show_r\nend program t\n",
        ["7", "2"]
    };

    generic_single_impl_alias => {
        "program t\ntype :: Box\ninteger :: w = 4\ncontains\nprocedure :: area_impl\ngeneric :: area => area_impl\nend type Box\ntype(Box) :: b\nprint *, b%area()\ncontains\ninteger function area_impl(self) result(a)\nclass(Box), intent(in) :: self\na = self%w * self%w\nend function area_impl\nend program t\n",
        ["16"]
    };

    generic_add_integers => {
        "program t\ntype :: Adder\ncontains\nprocedure :: add_i\ngeneric :: add => add_i\nend type Adder\ntype(Adder) :: a\nprint *, a%add(3, 5)\ncontains\ninteger function add_i(self, x, y) result(r)\nclass(Adder), intent(in) :: self\ninteger, intent(in) :: x, y\nr = x + y\nend function add_i\nend program t\n",
        ["8"]
    };

    generic_push_subroutine_overload => {
        "program t\ntype :: Stack\ninteger :: top = 0\ncontains\nprocedure :: push_i\nprocedure :: push_r\ngeneric :: push => push_i, push_r\nend type Stack\ntype(Stack) :: s\ncall s%push(1)\ncall s%push(2.0)\nprint *, s%top\ncontains\nsubroutine push_i(self, v)\nclass(Stack), intent(inout) :: self\ninteger, intent(in) :: v\nself%top = self%top + v\nend subroutine push_i\nsubroutine push_r(self, v)\nclass(Stack), intent(inout) :: self\nreal, intent(in) :: v\nself%top = self%top + int(v)\nend subroutine push_r\nend program t\n",
        ["3"]
    };

    generic_module_bound_compare => {
        "module cmpmod\nimplicit none\ntype :: Cmp\ncontains\nprocedure :: eq_i\nprocedure :: eq_r\ngeneric :: eq => eq_i, eq_r\nend type Cmp\ncontains\nlogical function eq_i(self, a, b) result(r)\nclass(Cmp), intent(in) :: self\ninteger, intent(in) :: a, b\nr = a == b\nend function eq_i\nlogical function eq_r(self, a, b) result(r)\nclass(Cmp), intent(in) :: self\nreal, intent(in) :: a, b\nr = abs(a - b) < 1.0e-6\nend function eq_r\nend module cmpmod\nprogram t\nuse cmpmod\ntype(Cmp) :: c\nprint *, c%eq(4, 4)\nprint *, c%eq(1.0, 1.0)\nend program t\n",
        ["1", "1"]
    };

    // ── Deferred / abstract types ────────────────────────────────────

    deferred_rect_area_runtime => {
        "module shapes\nimplicit none\ntype, abstract :: Poly\ncontains\nprocedure(perim_iface), deferred :: perimeter\nend type Poly\nabstract interface\nfunction perim_iface(self) result(p)\nimport Poly\nclass(Poly), intent(in) :: self\nreal :: p\nend function perim_iface\nend interface\ntype, extends(Poly) :: Rect\nreal :: w, h\ncontains\nprocedure :: perimeter => rect_perim\nend type Rect\ncontains\nfunction rect_perim(self) result(p)\nclass(Rect), intent(in) :: self\nreal :: p\np = 2.0 * (self%w + self%h)\nend function rect_perim\nend module shapes\nprogram t\nuse shapes\ntype(Rect) :: r\nr%w = 3.0\nr%h = 4.0\nprint *, int(r%perimeter())\nend program t\n",
        ["14"]
    };

    deferred_triangle_area_runtime => {
        "module geom\nimplicit none\ntype, abstract :: Figure\ncontains\nprocedure(area_iface), deferred :: area\nend type Figure\nabstract interface\nfunction area_iface(self) result(a)\nimport Figure\nclass(Figure), intent(in) :: self\nreal :: a\nend function area_iface\nend interface\ntype, extends(Figure) :: Tri\nreal :: base, height\ncontains\nprocedure :: area => tri_area\nend type Tri\ncontains\nfunction tri_area(self) result(a)\nclass(Tri), intent(in) :: self\nreal :: a\na = 0.5 * self%base * self%height\nend function tri_area\nend module geom\nprogram t\nuse geom\ntype(Tri) :: t\nt%base = 6.0\nt%height = 4.0\nprint *, int(t%area())\nend program t\n",
        ["12"]
    };

    polymorphic_abstract_greet => {
        "module greetmod\nimplicit none\ntype, abstract :: Greeter\ncontains\nprocedure(msg_iface), deferred :: message\nend type Greeter\nabstract interface\nfunction msg_iface(self) result(s)\nimport Greeter\nclass(Greeter), intent(in) :: self\ncharacter(len=16) :: s\nend function msg_iface\nend interface\ntype, extends(Greeter) :: Hello\ncontains\nprocedure :: message => hello_msg\nend type Hello\ncontains\nfunction hello_msg(self) result(s)\nclass(Hello), intent(in) :: self\ncharacter(len=16) :: s\ns = 'hello'\nend function hello_msg\nend module greetmod\nprogram t\nuse greetmod\ntype(Hello) :: h\nprint *, trim(h%message())\nend program t\n",
        ["hello"]
    };

    class_allocatable_abstract_dispatch => {
        "module animals\nimplicit none\ntype, abstract :: Animal\ninteger :: legs = 0\ncontains\nprocedure(legs_iface), deferred :: count_legs\nend type Animal\nabstract interface\nfunction legs_iface(self) result(n)\nimport Animal\nclass(Animal), intent(in) :: self\ninteger :: n\nend function legs_iface\nend interface\ntype, extends(Animal) :: Spider\ncontains\nprocedure :: count_legs => spider_legs\nend type Spider\ncontains\nfunction spider_legs(self) result(n)\nclass(Spider), intent(in) :: self\nn = 8\nend function spider_legs\nend module animals\nprogram t\nuse animals\nclass(Animal), allocatable :: a\nallocate(Spider :: a)\nprint *, a%count_legs()\nend program t\n",
        ["8"]
    };

    extends_type_of_deferred_child => {
        "module bases\nimplicit none\ntype :: Root\ninteger :: id = 1\nend type Root\ntype, abstract, extends(Root) :: Node\ncontains\nprocedure(val_iface), deferred :: value\nend type Node\nabstract interface\nfunction val_iface(self) result(v)\nimport Node\nclass(Node), intent(in) :: self\ninteger :: v\nend function val_iface\nend interface\ntype, extends(Node) :: Leaf\ninteger :: payload = 9\ncontains\nprocedure :: value => leaf_val\nend type Leaf\ncontains\nfunction leaf_val(self) result(v)\nclass(Leaf), intent(in) :: self\nv = self%payload\nend function leaf_val\nend module bases\nprogram t\nuse bases\ntype(Root) :: r\ntype(Leaf) :: leaf\nprint *, extends_type_of(leaf, r)\nprint *, leaf%value()\nend program t\n",
        ["1", "9"]
    };

    // ── Allocatable components ───────────────────────────────────────

    alloc_comp_int_vector_sum => {
        "program t\ntype :: Vec\ninteger, allocatable :: data(:)\nend type Vec\ntype(Vec) :: v\nv%data = [2, 4, 6]\nprint *, sum(v%data)\nend program t\n",
        ["12"]
    };

    alloc_comp_real_matrix_corner => {
        "program t\ntype :: Grid\nreal, allocatable :: cells(:,:)\nend type Grid\ntype(Grid) :: g\nallocate(g%cells(2, 2))\ng%cells = reshape([1.0, 2.0, 3.0, 4.0], [2, 2])\nprint *, int(g%cells(2, 1))\nend program t\n",
        ["3"]
    };

    alloc_comp_nested_container => {
        "program t\ntype :: Inner\ninteger :: key = 0\nend type Inner\ntype :: Outer\ninteger :: id = 1\ntype(Inner), allocatable :: payload(:)\nend type Outer\ntype(Outer) :: o\nallocate(o%payload(2))\no%payload(1)%key = 7\no%payload(2)%key = 3\nprint *, o%payload(1)%key + o%payload(2)%key\nend program t\n",
        ["10"]
    };

    alloc_comp_character_label => {
        "program t\ntype :: Tag\ncharacter(len=6), allocatable :: name\nend type Tag\ntype(Tag) :: t\nallocate(t%name)\nt%name = 'f2003'\nprint *, trim(t%name)\nend program t\n",
        ["f2003"]
    };

    alloc_comp_explicit_allocate_size => {
        "program t\ntype :: Buffer\ninteger, allocatable :: slots(:)\nend type Buffer\ntype(Buffer) :: b\nallocate(b%slots(4))\nb%slots = [10, 20, 30, 40]\nprint *, b%slots(3)\nend program t\n",
        ["30"]
    };

    alloc_comp_move_alloc_between_fields => {
        "program t\ntype :: Pair\ninteger, allocatable :: left(:), right(:)\nend type Pair\ntype(Pair) :: p\nallocate(p%left(2))\np%left = [5, 6]\ncall move_alloc(p%left, p%right)\nprint *, p%right(1)\nprint *, allocated(p%left)\nend program t\n",
        ["5", "0"]
    };

    alloc_comp_logical_flag_value => {
        "program t\ntype :: Flags\nlogical, allocatable :: bits(:)\nend type Flags\ntype(Flags) :: f\nf%bits = [.true., .false., .true.]\nprint *, count(f%bits)\nend program t\n",
        ["2"]
    };

    alloc_comp_module_derived_field => {
        "module nodes\nimplicit none\ntype :: Node\ninteger :: value = 0\nend type Node\ntype :: List\ntype(Node), allocatable :: items(:)\nend type List\nend module nodes\nprogram t\nuse nodes\ntype(List) :: lst\nallocate(lst%items(1))\nlst%items(1)%value = 42\nprint *, lst%items(1)%value\nend program t\n",
        ["42"]
    };

    // ── Stream I/O ───────────────────────────────────────────────────

    stream_scratch_write_read_ints => {
        "program t\ninteger :: a, b\nopen(10, status='scratch', access='stream', form='unformatted')\nwrite(10) 11, 22\nrewind(10)\nread(10) a, b\nclose(10)\nprint *, a + b\nend program t\n",
        ["33"]
    };

    stream_unformatted_real_pair => {
        "program t\nreal :: x, y\nopen(11, status='scratch', access='stream', form='unformatted')\nwrite(11) 1.5, 2.5\nrewind(11)\nread(11) x, y\nclose(11)\nprint *, int(x + y)\nend program t\n",
        ["4"]
    };

    stream_rewind_preserves_first_value => {
        "program t\ninteger :: first, second\nopen(12, status='scratch', access='stream', form='unformatted')\nwrite(12) 100\nwrite(12) 200\nrewind(12)\nread(12) first\nread(12) second\nclose(12)\nprint *, first\nprint *, second\nend program t\n",
        ["100", "200"]
    };

    stream_write_logical_scalar => {
        "program t\nlogical :: flag\nopen(13, status='scratch', access='stream', form='unformatted')\nwrite(13) .true.\nrewind(13)\nread(13) flag\nclose(13)\nprint *, flag\nend program t\n",
        ["1"]
    };

    stream_module_helper_roundtrip => {
        "module iostream\nimplicit none\ncontains\nsubroutine write_pair(u, a, b)\ninteger, intent(in) :: u, a, b\nwrite(u) a, b\nend subroutine write_pair\nsubroutine read_pair(u, a, b)\ninteger, intent(in) :: u\ninteger, intent(out) :: a, b\nread(u) a, b\nend subroutine read_pair\nend module iostream\nprogram t\nuse iostream\ninteger :: x, y\nopen(14, status='scratch', access='stream', form='unformatted')\ncall write_pair(14, 8, 9)\nrewind(14)\ncall read_pair(14, x, y)\nclose(14)\nprint *, x * y\nend program t\n",
        ["72"]
    };
}

// ── Compile-only F2003 constructs ───────────────────────────────────

#[test]
fn compile_tbp_private_binding_in_module() {
    compile_ok(
        r#"
module secrets
    implicit none
    type :: Vault
        integer :: code = 0
    contains
        procedure, private :: seal
        procedure :: open => unlock
    end type Vault
contains
    subroutine seal(self, c)
        class(Vault), intent(inout) :: self
        integer, intent(in) :: c
        self%code = c
    end subroutine seal
    subroutine unlock(self)
        class(Vault), intent(inout) :: self
        self%code = 0
    end subroutine unlock
end module secrets

program t
    use secrets
    type(Vault) :: v
    call v%open()
    print *, v%code
end program t
"#,
    );
}

#[test]
fn compile_tbp_non_overridable_binding() {
    compile_ok(
        r#"
program t
    type :: Fixed
        integer :: n = 1
    contains
        procedure, non_overridable :: id
    end type Fixed
    type(Fixed) :: f
    print *, f%id()
contains
    function id(self) result(v)
        class(Fixed), intent(in) :: self
        integer :: v
        v = self%n
    end function id
end program t
"#,
    );
}

#[test]
fn compile_generic_assignment_binding() {
    compile_ok(
        r#"
program t
    type :: Bag
        integer, allocatable :: items(:)
    contains
        procedure :: assign_from
        generic :: assignment(=) => assign_from
    end type Bag
    type(Bag) :: a, b
    a%items = [1, 2]
    b = a
    print *, size(b%items)
contains
    subroutine assign_from(lhs, rhs)
        class(Bag), intent(out) :: lhs
        type(Bag), intent(in) :: rhs
        lhs%items = rhs%items
    end subroutine assign_from
end program t
"#,
    );
}

#[test]
fn compile_generic_read_write_bindings() {
    compile_ok(
        r#"
program t
    type :: Pair
        integer :: a, b
    contains
        procedure :: read_pair
        procedure :: write_pair
        generic :: read(formatted) => read_pair
        generic :: write(formatted) => write_pair
    end type Pair
    type(Pair) :: p
    p%a = 1
    p%b = 2
    print *, p%a + p%b
contains
    subroutine read_pair(self, unit, iostat)
        class(Pair), intent(out) :: self
        integer, intent(in) :: unit
        integer, intent(out), optional :: iostat
        read(unit, *, iostat=iostat) self%a, self%b
    end subroutine read_pair
    subroutine write_pair(self, unit, iostat)
        class(Pair), intent(in) :: self
        integer, intent(in) :: unit
        integer, intent(out), optional :: iostat
        write(unit, *, iostat=iostat) self%a, self%b
    end subroutine write_pair
end program t
"#,
    );
}

#[test]
fn compile_deferred_two_abstract_procedures() {
    compile_ok(
        r#"
module iface2
    implicit none
    type, abstract :: Expr
    contains
        procedure(eval_iface), deferred :: eval
        procedure(arity_iface), deferred :: arity
    end type Expr

    abstract interface
        integer function eval_iface(self) result(v)
            import Expr
            class(Expr), intent(in) :: self
        end function eval_iface
        integer function arity_iface(self) result(n)
            import Expr
            class(Expr), intent(in) :: self
        end function arity_iface
    end interface
end module iface2

program t
    use iface2
    print *, "ok"
end program t
"#,
    );
}

#[test]
fn compile_abstract_extends_concrete_parent() {
    compile_ok(
        r#"
module hier
    implicit none
    type :: Entity
        integer :: uid = 0
    end type Entity
    type, abstract, extends(Entity) :: Drawable
    contains
        procedure(draw_iface), deferred :: draw
    end type Drawable

    abstract interface
        subroutine draw_iface(self)
            import Drawable
            class(Drawable), intent(in) :: self
        end subroutine draw_iface
    end interface
end module hier

program t
    use hier
    type(Entity) :: e
    e%uid = 7
    print *, e%uid
end program t
"#,
    );
}

#[test]
fn compile_alloc_comp_deferred_shape_rank() {
    compile_ok(
        r#"
program t
    type :: Flex
        integer, allocatable :: buf(:)
    end type Flex
    type(Flex) :: f
    allocate(f%buf(0:2))
    f%buf = [1, 2, 3]
    print *, f%buf(0) + f%buf(2)
end program t
"#,
    );
}

#[test]
fn compile_alloc_comp_pointer_component_coexist() {
    compile_ok(
        r#"
program t
    type :: Mix
        integer, allocatable :: owned(:)
        integer, pointer :: view(:) => null()
    end type Mix
    type(Mix) :: m
    allocate(m%owned(2))
    m%owned = [4, 5]
    m%view => m%owned
    print *, m%view(1)
end program t
"#,
    );
}

#[test]
fn compile_alloc_comp_default_init_in_type() {
    compile_ok(
        r#"
program t
    type :: Defaults
        integer, allocatable :: vals(:)
    end type Defaults
    type(Defaults) :: d
    d%vals = [9]
    print *, d%vals(1)
end program t
"#,
    );
}

#[test]
fn compile_flush_scratch_unit_after_write() {
    compile_ok(
        r#"
program t
    integer :: u = 10
    open(u, status='scratch')
    write(u, *) 42
    flush(u)
    close(u)
    print *, 'done'
end program t
"#,
    );
}

#[test]
fn compile_flush_output_unit_no_arg() {
    compile_ok(
        r#"
program t
    use iso_fortran_env, only: output_unit
    write(output_unit, *) 'line'
    flush(output_unit)
end program t
"#,
    );
}

#[test]
fn compile_flush_after_namelist_write() {
    compile_ok(
        r#"
program t
    integer :: n = 3
    real :: x = 1.5
    open(20, status='scratch')
    write(20, nml=cfg)
    flush(20)
    close(20)
    print *, n
contains
    namelist /cfg/ n, x
end program t
"#,
    );
}

#[test]
fn compile_stream_access_replace_file() {
    compile_ok(
        r#"
program t
    integer :: v
    open(30, file='tmp_stream.bin', access='stream', form='unformatted', status='replace')
    write(30) 77
    rewind(30)
    read(30) v
    close(30, status='delete')
    print *, v
end program t
"#,
    );
}

#[test]
fn compile_stream_position_inquire() {
    compile_ok(
        r#"
program t
    integer :: pos
    open(31, status='scratch', access='stream', form='unformatted')
    write(31) 1, 2, 3
    inquire(31, pos=pos)
    close(31)
    print *, pos
end program t
"#,
    );
}

#[test]
fn compile_stream_mixed_type_sequential() {
    compile_ok(
        r#"
program t
    integer :: i
    real :: r
    character(len=4) :: tag
    open(32, status='scratch', access='stream', form='unformatted')
    write(32) 5, 2.5
    rewind(32)
    read(32) i, r
    close(32)
    print *, i, int(r)
end program t
"#,
    );
}

#[test]
fn compile_polymorphic_select_type_deferred_child() {
    compile_ok(
        r#"
module poly
    implicit none
    type, abstract :: Op
    contains
        procedure(run_iface), deferred :: run
    end type Op
    abstract interface
        integer function run_iface(self) result(v)
            import Op
            class(Op), intent(in) :: self
        end function run_iface
    end interface
    type, extends(Op) :: Inc
        integer :: step = 1
    contains
        procedure :: run => inc_run
    end type Inc
contains
    function inc_run(self) result(v)
        class(Inc), intent(in) :: self
        v = self%step
    end function inc_run
end module poly

program t
    use poly
    class(Op), allocatable :: job
    allocate(Inc :: job)
    select type(job)
    type is (Inc)
        print *, job%run()
  class default
        print *, 0
    end select
end program t
"#,
    );
}

#[test]
fn compile_class_pointer_polymorphic_component() {
    compile_ok(
        r#"
program t
    type :: Base
        integer :: tag = 1
    end type Base
    type, extends(Base) :: Ext
        integer :: extra = 9
    end type Ext
    type :: Holder
        class(Base), pointer :: item => null()
    end type Holder
    type(Holder) :: h
    type(Ext), target :: e
    h%item => e
    print *, h%item%tag
end program t
"#,
    );
}

#[test]
fn compile_allocatable_polymorphic_array_component() {
    compile_ok(
        r#"
program t
    type :: Part
        integer :: id = 0
    end type Part
    type :: Assembly
        class(Part), allocatable :: parts(:)
    end type Assembly
    type(Assembly) :: a
    allocate(Part :: a%parts(2))
    a%parts(1)%id = 3
  a%parts(2)%id = 4
    print *, a%parts(1)%id + a%parts(2)%id
end program t
"#,
    );
}
