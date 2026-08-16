//! Extended pointer and allocatable coverage: allocate, deallocate, associated,
//! nullify, pointer assignment, allocatable assignment, and move_alloc.
//! Distinct from `test_pointers.rs` (basic declarations, targets, procedure pointers).

use super::helpers::compile_ok;

fortran_cases! {
    // ── ALLOCATE / DEALLOCATE / ALLOCATED ────────────────────────────

    alloc_int_five_element_sum => {
        "program t\ninteger, allocatable :: v(:)\nallocate(v(5))\nv = [(i, i = 1, 5)]\nprint *, sum(v)\ndeallocate(v)\nend program t\n",
        ["15"]
    };

    alloc_real_three_first_element => {
        "program t\nreal, allocatable :: r(:)\nallocate(r(3))\nr = [1.5, 2.5, 3.5]\nprint *, r(1)\ndeallocate(r)\nend program t\n",
        ["1.5"]
    };

    alloc_logical_scalar_value => {
        "program t\nlogical, allocatable :: flag\nallocate(flag)\nflag = .true.\nprint *, flag\ndeallocate(flag)\nend program t\n",
        ["true"]
    };

    alloc_char_two_element_trim => {
        "program t\ncharacter(len=4), allocatable :: words(:)\nallocate(words(2))\nwords(1) = 'ab'\nwords(2) = 'cd'\nprint *, trim(words(1))\nprint *, len_trim(words(2))\ndeallocate(words)\nend program t\n",
        ["ab", "2"]
    };

    alloc_2d_three_by_two_sum => {
        "program t\ninteger, allocatable :: m(:,:)\nallocate(m(3, 2))\nm = 1\nprint *, sum(m)\nprint *, size(m)\ndeallocate(m)\nend program t\n",
        ["6", "6"]
    };

    allocated_false_before_any_allocate => {
        "program t\ninteger, allocatable :: buf(:)\nprint *, allocated(buf)\nend program t\n",
        ["false"]
    };

    allocated_true_immediately_after_allocate => {
        "program t\ninteger, allocatable :: buf(:)\nallocate(buf(4))\nprint *, allocated(buf)\ndeallocate(buf)\nend program t\n",
        ["true"]
    };

    deallocate_then_allocated_false => {
        "program t\nreal, allocatable :: data(:)\nallocate(data(2))\ndata = [3.0, 4.0]\ndeallocate(data)\nprint *, allocated(data)\nend program t\n",
        ["false"]
    };

    deallocate_two_int_arrays_together => {
        "program t\ninteger, allocatable :: a(:), b(:)\nallocate(a(2), b(3))\na = [10, 20]\nb = [1, 2, 3]\nprint *, sum(a)\nprint *, sum(b)\ndeallocate(a, b)\nprint *, allocated(a)\nprint *, allocated(b)\nend program t\n",
        ["30", "6", "false", "false"]
    };

    allocate_stat_reports_success => {
        "program t\ninteger, allocatable :: v(:)\ninteger :: ierr\nallocate(v(3), stat=ierr)\nprint *, ierr\nprint *, size(v)\ndeallocate(v)\nend program t\n",
        ["0", "3"]
    };

    alloc_scalar_integer_value => {
        "program t\ninteger, allocatable :: n\nallocate(n)\nn = 99\nprint *, n\ndeallocate(n)\nend program t\n",
        ["99"]
    };

    alloc_3d_shape_product => {
        "program t\ninteger, allocatable :: cube(:,:,:)\nallocate(cube(2, 3, 4))\nprint *, size(cube)\nprint *, size(cube, 1)\nprint *, size(cube, 2)\nprint *, size(cube, 3)\ndeallocate(cube)\nend program t\n",
        ["24", "2", "3", "4"]
    };

    deallocate_then_reallocate_same_variable => {
        "program t\ninteger, allocatable :: v(:)\nallocate(v(2))\nv = [5, 6]\ndeallocate(v)\nallocate(v(4))\nv = [(i, i = 1, 4)]\nprint *, sum(v)\ndeallocate(v)\nend program t\n",
        ["10"]
    };

    // ── Allocatable assignment (=) ───────────────────────────────────

    alloc_assign_literal_three_ints => {
        "program t\ninteger, allocatable :: v(:)\nv = [4, 5, 6]\nprint *, v(2)\nprint *, size(v)\nend program t\n",
        ["5", "3"]
    };

    alloc_copy_between_two_arrays => {
        "program t\ninteger, allocatable :: src(:), dst(:)\nsrc = [7, 8, 9]\ndst = src\nprint *, dst(1)\nprint *, dst(3)\nend program t\n",
        ["7", "9"]
    };

    alloc_real_assign_literal_row => {
        "program t\nreal, allocatable :: row(:)\nrow = [0.5, 1.5, 2.5]\nprint *, row(2)\nprint *, size(row)\nend program t\n",
        ["1.5", "3"]
    };

    alloc_reassign_grows_array_size => {
        "program t\ninteger, allocatable :: items(:)\nitems = [1, 2]\nprint *, size(items)\nitems = [10, 20, 30, 40]\nprint *, size(items)\nprint *, items(4)\nend program t\n",
        ["2", "4", "40"]
    };

    alloc_2d_assign_literal_matrix => {
        "program t\ninteger, allocatable :: grid(:,:)\ngrid = reshape([1, 2, 3, 4], [2, 2])\nprint *, grid(2, 1)\nprint *, sum(grid)\nend program t\n",
        ["2", "10"]
    };

    alloc_derived_field_via_assignment => {
        "program t\ntype :: Bag\ninteger, allocatable :: items(:)\nend type Bag\ntype(Bag) :: box\nbox%items = [2, 4, 6]\nprint *, box%items(2)\nprint *, sum(box%items)\nend program t\n",
        ["4", "12"]
    };

    realloc_via_assignment_changes_length => {
        "program t\ninteger, allocatable :: seq(:)\nseq = [1]\nprint *, size(seq)\nseq = [1, 2, 3, 4, 5]\nprint *, size(seq)\nprint *, seq(5)\nend program t\n",
        ["1", "5", "5"]
    };

    // ── MOVE_ALLOC ─────────────────────────────────────────────────

    move_alloc_preserves_first_element => {
        "program t\ninteger, allocatable :: from(:), to(:)\nallocate(from(3))\nfrom = [11, 22, 33]\ncall move_alloc(from, to)\nprint *, to(1)\nprint *, size(to)\nend program t\n",
        ["11", "3"]
    };

    move_alloc_source_becomes_unallocated => {
        "program t\ninteger, allocatable :: src(:), dst(:)\nallocate(src(2))\nsrc = [5, 6]\ncall move_alloc(src, dst)\nprint *, allocated(src)\nprint *, allocated(dst)\nend program t\n",
        ["false", "true"]
    };

    move_alloc_real_vectors_sum => {
        "program t\nreal, allocatable :: a(:), b(:)\nallocate(a(4))\na = [1.0, 2.0, 3.0, 4.0]\ncall move_alloc(a, b)\nprint *, sum(b)\nprint *, allocated(a)\nend program t\n",
        ["10", "false"]
    };

    move_alloc_replaces_prior_destination => {
        "program t\ninteger, allocatable :: fresh(:), old(:)\nallocate(fresh(2))\nfresh = [100, 200]\nold = [9]\ncall move_alloc(fresh, old)\nprint *, old(1)\nprint *, size(old)\nend program t\n",
        ["100", "2"]
    };

    move_alloc_char_array_first_word => {
        "program t\ncharacter(len=3), allocatable :: a(:), b(:)\nallocate(a(2))\na(1) = 'xyz'\na(2) = 'uvw'\ncall move_alloc(a, b)\nprint *, trim(b(1))\nprint *, allocated(a)\nend program t\n",
        ["xyz", "false"]
    };

    // ── ASSOCIATED ─────────────────────────────────────────────────

    associated_unassociated_pointer_is_false => {
        "program t\ninteger, pointer :: p => null()\nprint *, associated(p)\nend program t\n",
        ["false"]
    };

    associated_after_pointer_assignment_true => {
        "program t\ninteger, target :: host = 17\ninteger, pointer :: view\nview => host\nprint *, associated(view)\nend program t\n",
        ["true"]
    };

    associated_with_matching_target_true => {
        "program t\ninteger, target :: x = 3, y = 4\ninteger, pointer :: link\nlink => x\nprint *, associated(link, x)\nprint *, associated(link, y)\nend program t\n",
        ["true", "false"]
    };

    associated_after_nullify_is_false => {
        "program t\ninteger, target :: val = 8\ninteger, pointer :: link\nlink => val\nnullify(link)\nprint *, associated(link)\nend program t\n",
        ["false"]
    };

    associated_derived_type_field_initially_false => {
        "program t\ntype :: Node\ninteger :: id\ninteger, pointer :: child => null()\nend type Node\ntype(Node) :: n\nn%id = 1\nprint *, associated(n%child)\nend program t\n",
        ["false"]
    };

    // ── NULLIFY ────────────────────────────────────────────────────

    nullify_then_reassociate_pointer => {
        "program t\ninteger, target :: a = 1, b = 2\ninteger, pointer :: p\np => a\nnullify(p)\nprint *, associated(p)\np => b\nprint *, p\nend program t\n",
        ["false", "2"]
    };

    nullify_pair_leaves_both_unassociated => {
        "program t\ninteger, target :: u = 1, v = 2\ninteger, pointer :: p, q\np => u\nq => v\nnullify(p, q)\nprint *, associated(p)\nprint *, associated(q)\nend program t\n",
        ["false", "false"]
    };

    // ── Pointer assignment (=>) ────────────────────────────────────

    pointer_read_scalar_through_association => {
        "program t\ninteger, target :: base = 42\ninteger, pointer :: alias\nalias => base\nprint *, alias\nend program t\n",
        ["42"]
    };

    pointer_write_updates_target_storage => {
        "program t\ninteger, target :: base = 1\ninteger, pointer :: alias\nalias => base\nalias = 50\nprint *, base\nend program t\n",
        ["50"]
    };

    pointer_array_third_element => {
        "program t\ninteger, target :: data(5) = [3, 6, 9, 12, 15]\ninteger, pointer :: slice(:)\nslice => data\nprint *, slice(3)\nend program t\n",
        ["9"]
    };

    pointer_assign_from_another_pointer => {
        "program t\ninteger, target :: val = 77\ninteger, pointer :: first, second\nfirst => val\nsecond => first\nprint *, second\nend program t\n",
        ["77"]
    };

    pointer_2d_matrix_center => {
        "program t\ninteger, target :: mat(2, 2)\ninteger, pointer :: view(:,:)\nmat = reshape([1, 2, 3, 4], [2, 2])\nview => mat\nprint *, view(2, 1)\nend program t\n",
        ["2"]
    };

    pointer_target_section_first_element => {
        "program t\ninteger, target :: series(6) = [10, 20, 30, 40, 50, 60]\ninteger, pointer :: window(:)\nwindow => series(2:4)\nprint *, window(1)\nprint *, size(window)\nend program t\n",
        ["20", "3"]
    };

    pointer_chain_across_two_targets => {
        "program t\ninteger, target :: left = 5, right = 15\ninteger, pointer :: hop\nhop => left\nprint *, hop\nhop => right\nprint *, hop\nend program t\n",
        ["5", "15"]
    };

    pointer_derived_next_link_value => {
        "program t\ntype :: Link\ninteger :: payload\ntype(Link), pointer :: nxt => null()\nend type Link\ntype(Link), target :: head, tail\nhead%payload = 1\ntail%payload = 2\nhead%nxt => tail\nprint *, head%nxt%payload\nend program t\n",
        ["2"]
    };
}

// ── Compile-only pointer / allocatable shapes ─────────────────────

#[test]
fn compile_pointer_target_in_subroutine() {
    compile_ok(
        r#"
subroutine bind_ptr(host, view)
    integer, target, intent(inout) :: host
    integer, pointer, intent(out) :: view
    view => host
end subroutine bind_ptr

program t
    integer, target :: x = 6
    integer, pointer :: p
    call bind_ptr(x, p)
    print *, p
end program t
"#,
    );
}

#[test]
fn compile_allocatable_intent_out_subroutine() {
    compile_ok(
        r#"
subroutine make_buffer(buf, n)
    integer, intent(in) :: n
    integer, allocatable, intent(out) :: buf(:)
    allocate(buf(n))
    buf = [(i, i = 1, n)]
end subroutine make_buffer

program t
    integer, allocatable :: data(:)
    call make_buffer(data, 3)
    print *, sum(data)
end program t
"#,
    );
}

#[test]
fn compile_move_alloc_inside_subroutine() {
    compile_ok(
        r#"
subroutine transfer_storage(from, to)
    integer, allocatable, intent(inout) :: from(:)
    integer, allocatable, intent(inout) :: to(:)
    call move_alloc(from, to)
end subroutine transfer_storage

program t
    integer, allocatable :: a(:), b(:)
    allocate(a(2))
    a = [3, 4]
    call transfer_storage(a, b)
    print *, b(2)
    print *, allocated(a)
end program t
"#,
    );
}

#[test]
fn compile_pointer_deferred_shape_dummy() {
    compile_ok(
        r#"
subroutine first_elt(vec, out)
    integer, pointer, intent(in) :: vec(:)
    integer, intent(out) :: out
    out = vec(1)
end subroutine first_elt

program t
    integer, target :: arr(4) = [8, 6, 4, 2]
    integer, pointer :: p(:)
    integer :: head
    p => arr
    call first_elt(p, head)
    print *, head
end program t
"#,
    );
}

#[test]
fn compile_allocatable_function_result() {
    compile_ok(
        r#"
function doubled(n) result(out)
    integer, intent(in) :: n
    integer, allocatable :: out(:)
    allocate(out(n))
    out = [(2 * i, i = 1, n)]
end function doubled

program t
    integer, allocatable :: v(:)
    v = doubled(3)
    print *, v(3)
end program t
"#,
    );
}

#[test]
fn compile_nested_type_allocatable_allocate() {
    compile_ok(
        r#"
type :: Inner
    real, allocatable :: coeffs(:)
end type Inner

type :: Outer
    type(Inner) :: layer
end type Outer

program t
    type(Outer) :: obj
    allocate(obj%layer%coeffs(3))
    obj%layer%coeffs = [1.0, 2.0, 3.0]
    print *, obj%layer%coeffs(2)
    deallocate(obj%layer%coeffs)
end program t
"#,
    );
}

#[test]
fn compile_allocate_source_from_expression() {
    compile_ok(
        r#"
program t
    integer, allocatable :: base(:), copy(:)
    base = [2, 4, 6, 8]
    allocate(copy, source=base + 1)
    print *, copy(1)
    print *, copy(4)
    deallocate(base, copy)
end program t
"#,
    );
}

#[test]
fn compile_allocate_mold_preserves_rank_only() {
    compile_ok(
        r#"
program t
    integer, allocatable :: pattern(:,:), blank(:,:)
    allocate(pattern(2, 3))
    pattern = 0
    allocate(blank, mold=pattern)
    print *, size(blank, 1)
    print *, size(blank, 2)
    deallocate(pattern, blank)
end program t
"#,
    );
}

#[test]
fn compile_nullify_reassociate_multiple_pointers() {
    compile_ok(
        r#"
program t
    integer, target :: s = 1, t = 2, u = 3
    integer, pointer :: p => null(), q => null(), r => null()
    p => s
    q => t
    r => u
    nullify(p, q, r)
    p => t
    r => s
    print *, p
    print *, r
    print *, associated(q)
end program t
"#,
    );
}

#[test]
fn compile_pointer_and_allocatable_in_module() {
    compile_ok(
        r#"
module storage
    implicit none
    type :: Slot
        integer, pointer :: view => null()
        integer, allocatable :: owned(:)
    end type Slot
contains
    subroutine attach(slot, target_arr)
        type(Slot), intent(inout) :: slot
        integer, target, intent(in) :: target_arr(:)
        slot%view => target_arr
        if (.not. allocated(slot%owned)) allocate(slot%owned(size(target_arr)))
        slot%owned = target_arr
    end subroutine attach
end module storage

program t
    use storage
    integer, target :: data(2) = [11, 22]
    type(Slot) :: box
    call attach(box, data)
    print *, box%view(2)
    print *, box%owned(1)
end program t
"#,
    );
}
