//! Extended BLOCK construct: local declarations, variable shadowing, nested blocks,
//! allocatable/pointer locals, and EXIT/CYCLE interaction with enclosing loops.
//! Distinct from `test_fortran2008.rs` (basic block) and `test_control_flow_extended.rs`
//! (two block END-boundary tests).

fortran_cases! {
    // ── Local integer declarations ───────────────────────────────────

    block_local_integer_init => {
        "program t\ninteger :: outer = 10\nblock\ninteger :: inner\ninner = outer + 5\nprint *, inner\nend block\nend program t\n",
        ["15"]
    };

    block_local_integer_shadows_outer => {
        "program t\ninteger :: x = 1\nblock\ninteger :: x\nx = 100\nprint *, x\nend block\nprint *, x\nend program t\n",
        ["100", "1"]
    };

    block_local_integer_uninitialized_set => {
        "program t\nblock\ninteger :: n\nn = 7\nprint *, n\nend block\nend program t\n",
        ["7"]
    };

    block_local_two_integers => {
        "program t\nblock\ninteger :: a, b\na = 3\nb = 4\nprint *, a + b\nend block\nend program t\n",
        ["7"]
    };

    block_local_integer_from_outer_expr => {
        "program t\ninteger :: base = 6\nblock\ninteger :: scaled\nscaled = base * 3\nprint *, scaled\nend block\nend program t\n",
        ["18"]
    };

    // ── Local real and character ─────────────────────────────────────

    block_local_real_computation => {
        "program t\nreal :: pi = 3.14159\nblock\nreal :: area\narea = pi * 4.0\nprint *, int(area)\nend block\nend program t\n",
        ["12"]
    };

    block_local_real_shadows_outer => {
        "program t\nreal :: r = 1.5\nblock\nreal :: r\nr = 9.0\nprint *, int(r)\nend block\nprint *, int(r)\nend program t\n",
        ["9", "1"]
    };

    block_local_character_string => {
        "program t\nblock\ncharacter(len=6) :: msg\nmsg = 'block'\nprint *, trim(msg)\nend block\nend program t\n",
        ["block"]
    };

    block_local_character_shadow => {
        "program t\ncharacter(len=4) :: s = 'outer'\nblock\ncharacter(len=4) :: s\ns = 'inner'\nprint *, trim(s)\nend block\nprint *, trim(s)\nend program t\n",
        ["inner", "outer"]
    };

    block_local_logical_flag => {
        "program t\nblock\nlogical :: ok\nok = .true.\nprint *, ok\nend block\nend program t\n",
        ["true"]
    };

    // ── Nested blocks ────────────────────────────────────────────────

    block_nested_two_levels => {
        "program t\ninteger :: a = 1\nblock\ninteger :: b\nb = a + 2\nblock\ninteger :: c\nc = b + 3\nprint *, c\nend block\nend block\nend program t\n",
        ["6"]
    };

    block_nested_three_levels_sum => {
        "program t\nblock\ninteger :: l1\nl1 = 1\nblock\ninteger :: l2\nl2 = l1 + 2\nblock\ninteger :: l3\nl3 = l2 + 3\nprint *, l3\nend block\nend block\nend block\nend program t\n",
        ["6"]
    };

    block_nested_shadow_at_each_level => {
        "program t\ninteger :: v = 0\nblock\ninteger :: v\nv = 10\nblock\ninteger :: v\nv = 20\nprint *, v\nend block\nprint *, v\nend block\nprint *, v\nend program t\n",
        ["20", "10", "0"]
    };

    block_nested_inner_only_sees_middle => {
        "program t\ninteger :: x = 5\nblock\ninteger :: y\ny = x + 1\nblock\ninteger :: z\nz = y + 2\nprint *, z\nend block\nprint *, y\nend block\nend program t\n",
        ["8", "6"]
    };

    block_nested_real_inner => {
        "program t\nblock\nreal :: outer_r\nouter_r = 2.0\nblock\nreal :: inner_r\ninner_r = outer_r * 3.0\nprint *, int(inner_r)\nend block\nend block\nend program t\n",
        ["6"]
    };

    // ── Modify outer (non-shadowed) variables ────────────────────────

    block_modify_outer_integer => {
        "program t\ninteger :: total = 0\nblock\ninteger :: addend\naddend = 7\ntotal = total + addend\nend block\nprint *, total\nend program t\n",
        ["7"]
    };

    block_modify_outer_in_loop => {
        "program t\ninteger :: sum = 0, i\nblock\ninteger :: term\ndo i = 1, 4\nterm = i\nsum = sum + term\nend do\nend block\nprint *, sum\nend program t\n",
        ["10"]
    };

    block_swap_via_temp => {
        "program t\ninteger :: a = 3, b = 9\nblock\ninteger :: tmp\ntmp = a\na = b\nb = tmp\nend block\nprint *, a\nprint *, b\nend program t\n",
        ["9", "3"]
    };

    block_increment_outer_counter => {
        "program t\ninteger :: n = 10\nblock\ninteger :: delta\ndelta = 5\nn = n + delta\nend block\nprint *, n\nend program t\n",
        ["15"]
    };

    // ── Block with allocatable locals ────────────────────────────────

    block_allocatable_integer_array => {
        "program t\nblock\ninteger, allocatable :: buf(:)\nallocate(buf(4))\nbuf = [1, 2, 3, 4]\nprint *, sum(buf)\ndeallocate(buf)\nend block\nend program t\n",
        ["10"]
    };

    block_allocatable_real_array => {
        "program t\nblock\nreal, allocatable :: vals(:)\nallocate(vals(3))\nvals = [1.0, 2.0, 3.0]\nprint *, int(sum(vals))\nend block\nend program t\n",
        ["6"]
    };

    block_allocatable_2d_array => {
        "program t\nblock\ninteger, allocatable :: m(:,:)\nallocate(m(2,2))\nm = reshape([1, 2, 3, 4], [2,2])\nprint *, m(2,1)\ndeallocate(m)\nend block\nend program t\n",
        ["3"]
    };

    block_allocatable_character => {
        "program t\nblock\ncharacter(len=:), allocatable :: s\ns = 'hello'\nprint *, len_trim(s)\nend block\nend program t\n",
        ["5"]
    };

    // ── Block with pointer locals ────────────────────────────────────

    block_pointer_to_outer_target => {
        "program t\ninteger, target :: host = 42\nblock\ninteger, pointer :: view\nview => host\nprint *, view\nend block\nend program t\n",
        ["42"]
    };

    block_pointer_modify_target => {
        "program t\ninteger, target :: val = 1\nblock\ninteger, pointer :: p\np => val\np = 99\nend block\nprint *, val\nend program t\n",
        ["99"]
    };

    block_pointer_reassign_in_block => {
        "program t\ninteger, target :: a = 5, b = 8\nblock\ninteger, pointer :: p\np => a\nprint *, p\np => b\nprint *, p\nend block\nend program t\n",
        ["5", "8"]
    };

    // ── Block with derived type locals ───────────────────────────────

    block_derived_type_local => {
        "program t\ntype :: Item\ninteger :: id\nend type Item\nblock\ntype(Item) :: it\nit%id = 42\nprint *, it%id\nend block\nend program t\n",
        ["42"]
    };

    block_derived_type_with_array_field => {
        "program t\ntype :: Bundle\ninteger :: data(3)\nend type Bundle\nblock\ntype(Bundle) :: b\nb%data = [2, 4, 6]\nprint *, sum(b%data)\nend block\nend program t\n",
        ["12"]
    };

    block_derived_type_shadow => {
        "program t\ntype :: Node\ninteger :: key = 0\nend type Node\ntype(Node) :: outer\nouter%key = 1\nblock\ntype(Node) :: outer\nouter%key = 99\nprint *, outer%key\nend block\nprint *, outer%key\nend program t\n",
        ["99", "1"]
    };

    // ── Block inside other constructs ────────────────────────────────

    block_inside_if_then => {
        "program t\ninteger :: x = 5\nif (x > 0) then\nblock\ninteger :: y\ny = x * 2\nprint *, y\nend block\nend if\nend program t\n",
        ["10"]
    };

    block_inside_if_else => {
        "program t\ninteger :: x = -3\nif (x >= 0) then\nprint *, 0\nelse\nblock\ninteger :: m\nm = abs(x)\nprint *, m\nend block\nend if\nend program t\n",
        ["3"]
    };

    block_inside_do_loop => {
        "program t\ninteger :: i, s\ns = 0\ndo i = 1, 3\nblock\ninteger :: sq\nsq = i * i\ns = s + sq\nend block\nend do\nprint *, s\nend program t\n",
        ["14"]
    };

    block_inside_select_case => {
        "program t\ninteger :: code = 2\nselect case (code)\ncase (1)\nprint *, 1\ncase (2)\nblock\ninteger :: v\nv = 20\nprint *, v\nend block\ncase default\nprint *, 0\nend select\nend program t\n",
        ["20"]
    };

    block_sequential_two_blocks => {
        "program t\nblock\ninteger :: a\na = 3\nprint *, a\nend block\nblock\ninteger :: b\nb = 7\nprint *, b\nend block\nend program t\n",
        ["3", "7"]
    };

    // ── Block with EXIT from enclosing loop ──────────────────────────

    block_exit_outer_labeled_do => {
        "program t\ninteger :: i\nouter: do i = 1, 10\nblock\nif (i == 4) exit outer\nprint *, i\nend block\nend do outer\nend program t\n",
        ["1", "2", "3"]
    };

    block_exit_at_first_iteration => {
        "program t\ninteger :: i, count\ncount = 0\nouter: do i = 1, 100\nblock\nif (i == 1) exit outer\ncount = count + 1\nend block\nend do outer\nprint *, count\nend program t\n",
        ["0"]
    };

    block_exit_after_three_prints => {
        "program t\ninteger :: k\nouter: do k = 1, 20\nblock\nif (k > 3) exit outer\nprint *, k\nend block\nend do outer\nend program t\n",
        ["1", "2", "3"]
    };

    block_exit_from_nested_do => {
        "program t\ninteger :: i, j\nouter: do i = 1, 5\ndo j = 1, 5\nblock\nif (j == 2) exit outer\nprint *, i * 10 + j\nend block\nend do\nend do outer\nend program t\n",
        ["11"]
    };

    // ── Block with CYCLE in enclosing loop ───────────────────────────

    block_cycle_skips_even => {
        "program t\ninteger :: i\ndo i = 1, 6\nblock\nif (mod(i, 2) == 0) cycle\nprint *, i\nend block\nend do\nend program t\n",
        ["1", "3", "5"]
    };

    block_cycle_labeled_outer => {
        "program t\ninteger :: i, j\nouter: do i = 1, 3\ndo j = 1, 3\nblock\nif (j == 2) cycle outer\nprint *, i * 10 + j\nend block\nend do\nend do outer\nend program t\n",
        ["11", "13", "31", "33"]
    };

    block_cycle_skip_multiples_of_three => {
        "program t\ninteger :: n\ndo n = 1, 9\nblock\nif (mod(n, 3) == 0) cycle\nprint *, n\nend block\nend do\nend program t\n",
        ["1", "2", "4", "5", "7", "8"]
    };

    // ── Block with complex and mixed types ───────────────────────────

    block_local_complex => {
        "program t\nblock\ncomplex :: z\nz = (3.0, 4.0)\nprint *, int(real(z) + aimag(z))\nend block\nend program t\n",
        ["7"]
    };

    block_local_parameter => {
        "program t\nblock\ninteger, parameter :: max = 100\nprint *, max\nend block\nend program t\n",
        ["100"]
    };

    block_local_kind_explicit => {
        "program t\nblock\ninteger(kind=8) :: big\nbig = 10000000000\nprint *, int(big / 1000000000)\nend block\nend program t\n",
        ["10"]
    };

    block_local_array_fixed_size => {
        "program t\nblock\ninteger :: arr(4)\narr = [1, 2, 3, 4]\nprint *, arr(3)\nend block\nend program t\n",
        ["3"]
    };

    block_local_array_sum => {
        "program t\nblock\ninteger :: data(5)\ndata = [10, 20, 30, 40, 50]\nprint *, sum(data)\nend block\nend program t\n",
        ["150"]
    };

    block_local_2d_array_access => {
        "program t\nblock\ninteger :: grid(2,2)\ngrid = reshape([1, 2, 3, 4], [2,2])\nprint *, grid(2,2)\nend block\nend program t\n",
        ["4"]
    };

    block_with_associate_inside => {
        "program t\ninteger :: x = 8\nblock\nassociate (y => x + 2)\nprint *, y\nend associate\nend block\nend program t\n",
        ["10"]
    };

    block_if_inside_block => {
        "program t\nblock\ninteger :: n\nn = 15\nif (n > 10) then\nprint *, 'big'\nelse\nprint *, 'small'\nend if\nend block\nend program t\n",
        ["big"]
    };

    block_do_inside_block_accumulate => {
        "program t\nblock\ninteger :: i, s\ns = 0\ndo i = 1, 5\ns = s + i\nend do\nprint *, s\nend block\nend program t\n",
        ["15"]
    };

    block_outer_unchanged_after_inner_shadow => {
        "program t\nreal :: temperature = 20.0\nblock\nreal :: temperature\ntemperature = 100.0\nprint *, int(temperature)\nend block\nprint *, int(temperature)\nend program t\n",
        ["100", "20"]
    };

    block_read_outer_write_inner => {
        "program t\ninteger :: limit = 5\nblock\ninteger :: i, total\ntotal = 0\ndo i = 1, limit\ntotal = total + i\nend do\nprint *, total\nend block\nend program t\n",
        ["15"]
    };

    block_nested_with_exit_inner_only => {
        "program t\ninteger :: i\nouter: do i = 1, 5\nblock\ninteger :: j\ndo j = 1, 3\nif (j == 2) exit\nprint *, i * 10 + j\nend do\nend block\nend do outer\nend program t\n",
        ["11", "21", "31", "41", "51"]
    };

    block_local_save_not_visible_outside => {
        "program t\nblock\ninteger :: secret\nsecret = 42\nprint *, secret\nend block\nprint *, 'done'\nend program t\n",
        ["42", "done"]
    };

    block_triple_nested_exit_value => {
        "program t\nblock\ninteger :: a\na = 1\nblock\ninteger :: b\nb = a + 1\nblock\ninteger :: c\nc = b + 1\nprint *, c\nend block\nend block\nend block\nend program t\n",
        ["3"]
    };
}
