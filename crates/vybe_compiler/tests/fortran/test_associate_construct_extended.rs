//! Extended ASSOCIATE construct: expression rename, array element/section,
//! derived-type components, nested associate, and associate inside blocks/loops.
//! Distinct from `test_modules_advanced.rs` (basic associate compile-only).

fortran_cases! {
    // ── Scalar variable rename ───────────────────────────────────────

    associate_scalar_integer_rename => {
        "program t\ninteger :: n = 42\nassociate (alias => n)\nprint *, alias\nend associate\nend program t\n",
        ["42"]
    };

    associate_scalar_real_rename => {
        "program t\nreal :: x = 3.5\nassociate (r => x)\nprint *, int(r * 2.0)\nend associate\nend program t\n",
        ["7"]
    };

    associate_scalar_character_rename => {
        "program t\ncharacter(len=6) :: word = 'fortran'\nassociate (w => word)\nprint *, trim(w)\nend associate\nend program t\n",
        ["fortran"]
    };

    associate_scalar_logical_rename => {
        "program t\nlogical :: flag = .true.\nassociate (f => flag)\nprint *, f\nend associate\nend program t\n",
        ["true"]
    };

    associate_scalar_modify_target => {
        "program t\ninteger :: count = 1\nassociate (c => count)\nc = c + 4\nend associate\nprint *, count\nend program t\n",
        ["5"]
    };

    // ── Expression associate (read-only) ─────────────────────────────

    associate_expr_sum_two_vars => {
        "program t\ninteger :: a = 10, b = 32\nassociate (total => a + b)\nprint *, total\nend associate\nend program t\n",
        ["42"]
    };

    associate_expr_product => {
        "program t\ninteger :: m = 6, n = 7\nassociate (prod => m * n)\nprint *, prod\nend associate\nend program t\n",
        ["42"]
    };

    associate_expr_sqrt_hypotenuse => {
        "program t\nreal :: x = 3.0, y = 4.0\nassociate (hyp => sqrt(x*x + y*y))\nprint *, int(hyp)\nend associate\nend program t\n",
        ["5"]
    };

    associate_expr_abs_negative => {
        "program t\ninteger :: v = -17\nassociate (mag => abs(v))\nprint *, mag\nend associate\nend program t\n",
        ["17"]
    };

    associate_expr_min_of_pair => {
        "program t\ninteger :: p = 8, q = 3\nassociate (lo => min(p, q))\nprint *, lo\nend associate\nend program t\n",
        ["3"]
    };

    associate_expr_max_of_pair => {
        "program t\ninteger :: p = 8, q = 3\nassociate (hi => max(p, q))\nprint *, hi\nend associate\nend program t\n",
        ["8"]
    };

    associate_expr_mod_remainder => {
        "program t\ninteger :: num = 17, den = 5\nassociate (rem => mod(num, den))\nprint *, rem\nend associate\nend program t\n",
        ["2"]
    };

    associate_expr_negation => {
        "program t\ninteger :: val = 9\nassociate (neg => -val)\nprint *, neg\nend associate\nend program t\n",
        ["-9"]
    };

    associate_expr_real_division => {
        "program t\nreal :: a = 7.0, b = 2.0\nassociate (q => a / b)\nprint *, int(q)\nend associate\nend program t\n",
        ["3"]
    };

    associate_expr_logical_and => {
        "program t\nlogical :: p = .true., q = .false.\nassociate (both => p .and. q)\nprint *, both\nend associate\nend program t\n",
        ["false"]
    };

    associate_expr_logical_or => {
        "program t\nlogical :: p = .true., q = .false.\nassociate (either => p .or. q)\nprint *, either\nend associate\nend program t\n",
        ["true"]
    };

    associate_expr_char_concat => {
        "program t\ncharacter(len=8) :: a = 'foo', b = 'bar'\nassociate (ab => trim(a) // trim(b))\nprint *, ab\nend associate\nend program t\n",
        ["foobar"]
    };

    // ── Array element and section ────────────────────────────────────

    associate_array_first_element => {
        "program t\ninteger :: a(5) = [11, 22, 33, 44, 55]\nassociate (head => a(1))\nprint *, head\nend associate\nend program t\n",
        ["11"]
    };

    associate_array_last_element => {
        "program t\ninteger :: a(5) = [11, 22, 33, 44, 55]\nassociate (tail => a(5))\nprint *, tail\nend associate\nend program t\n",
        ["55"]
    };

    associate_array_middle_element => {
        "program t\ninteger :: a(5) = [11, 22, 33, 44, 55]\nassociate (mid => a(3))\nprint *, mid\nend associate\nend program t\n",
        ["33"]
    };

    associate_array_element_modify => {
        "program t\ninteger :: a(4) = [1, 2, 3, 4]\nassociate (slot => a(2))\nslot = 99\nend associate\nprint *, a(2)\nend program t\n",
        ["99"]
    };

    associate_array_2d_element => {
        "program t\ninteger :: m(2,3) = reshape([1,2,3,4,5,6], [2,3])\nassociate (cell => m(2,1))\nprint *, cell\nend associate\nend program t\n",
        ["4"]
    };

    associate_array_2d_diagonal => {
        "program t\ninteger :: m(3,3)\nm = 0\nm(1,1) = 7\nm(2,2) = 8\nm(3,3) = 9\nassociate (d => m(2,2))\nprint *, d\nend associate\nend program t\n",
        ["8"]
    };

    associate_array_section_sum => {
        "program t\ninteger :: a(6) = [1, 2, 3, 4, 5, 6]\nassociate (slice => a(2:4))\nprint *, sum(slice)\nend associate\nend program t\n",
        ["9"]
    };

    associate_array_section_first => {
        "program t\ninteger :: a(5) = [10, 20, 30, 40, 50]\nassociate (slice => a(1:2))\nprint *, slice(2)\nend associate\nend program t\n",
        ["20"]
    };

    associate_array_whole_vector => {
        "program t\ninteger :: a(4) = [2, 4, 6, 8]\nassociate (vec => a)\nprint *, sum(vec)\nend associate\nend program t\n",
        ["20"]
    };

    associate_real_array_element => {
        "program t\nreal :: r(3) = [1.5, 2.5, 3.5]\nassociate (mid => r(2))\nprint *, int(mid * 2.0)\nend associate\nend program t\n",
        ["5"]
    };

    // ── Derived type components ──────────────────────────────────────

    associate_dtype_integer_field => {
        "program t\ntype :: Pair\ninteger :: x, y\nend type Pair\ntype(Pair) :: p\np%x = 12\np%y = 30\nassociate (first => p%x)\nprint *, first\nend associate\nend program t\n",
        ["12"]
    };

    associate_dtype_real_field => {
        "program t\ntype :: Point\nreal :: x, y\nend type Point\ntype(Point) :: pt\npt%x = 6.0\npt%y = 8.0\nassociate (abscissa => pt%x)\nprint *, int(abscissa)\nend associate\nend program t\n",
        ["6"]
    };

    associate_dtype_char_field => {
        "program t\ntype :: Label\ncharacter(len=5) :: text\nend type Label\ntype(Label) :: lbl\nlbl%text = 'hello'\nassociate (msg => lbl%text)\nprint *, trim(msg)\nend associate\nend program t\n",
        ["hello"]
    };

    associate_dtype_nested_component => {
        "program t\ntype :: Inner\ninteger :: val = 0\nend type Inner\ntype :: Outer\ntype(Inner) :: core\nend type Outer\ntype(Outer) :: o\no%core%val = 77\nassociate (v => o%core%val)\nprint *, v\nend associate\nend program t\n",
        ["77"]
    };

    associate_dtype_field_modify => {
        "program t\ntype :: Counter\ninteger :: n = 0\nend type Counter\ntype(Counter) :: c\nassociate (count => c%n)\ncount = 15\nend associate\nprint *, c%n\nend program t\n",
        ["15"]
    };

    associate_dtype_two_fields => {
        "program t\ntype :: Coord\ninteger :: x, y\nend type Coord\ntype(Coord) :: c\nc%x = 3\nc%y = 4\nassociate (px => c%x, py => c%y)\nprint *, px + py\nend associate\nend program t\n",
        ["7"]
    };

    associate_dtype_array_component_elem => {
        "program t\ntype :: Bag\ninteger :: items(3)\nend type Bag\ntype(Bag) :: b\nb%items = [5, 10, 15]\nassociate (second => b%items(2))\nprint *, second\nend associate\nend program t\n",
        ["10"]
    };

    // ── Nested and sequential associate ──────────────────────────────

    associate_nested_inner_expr => {
        "program t\ninteger :: base = 5\nassociate (outer => base * 2)\nassociate (inner => outer + 3)\nprint *, inner\nend associate\nend associate\nend program t\n",
        ["13"]
    };

    associate_nested_array_then_elem => {
        "program t\ninteger :: a(4) = [2, 4, 6, 8]\nassociate (vec => a)\nassociate (elem => vec(3))\nprint *, elem\nend associate\nend associate\nend program t\n",
        ["6"]
    };

    associate_sequential_two_blocks => {
        "program t\ninteger :: x = 3, y = 7\nassociate (a => x)\nprint *, a\nend associate\nassociate (b => y)\nprint *, b\nend associate\nend program t\n",
        ["3", "7"]
    };

    associate_nested_dtype_field => {
        "program t\ntype :: Node\ninteger :: key\nend type Node\ntype(Node) :: n\nn%key = 42\nassociate (item => n)\nassociate (k => item%key)\nprint *, k\nend associate\nend associate\nend program t\n",
        ["42"]
    };

    // ── Associate inside other constructs ────────────────────────────

    associate_inside_block => {
        "program t\ninteger :: x = 8\nblock\ninteger :: y\ny = x + 2\nassociate (z => y * 3)\nprint *, z\nend associate\nend block\nend program t\n",
        ["30"]
    };

    associate_inside_if_then => {
        "program t\ninteger :: n = 6\nif (n > 0) then\nassociate (sq => n * n)\nprint *, sq\nend associate\nend if\nend program t\n",
        ["36"]
    };

    associate_inside_if_else => {
        "program t\ninteger :: n = -4\nif (n >= 0) then\nprint *, 0\nelse\nassociate (mag => abs(n))\nprint *, mag\nend associate\nend if\nend program t\n",
        ["4"]
    };

    associate_inside_do_loop => {
        "program t\ninteger :: i, s\ns = 0\ndo i = 1, 3\nassociate (term => i * i)\ns = s + term\nend associate\nend do\nprint *, s\nend program t\n",
        ["14"]
    };

    associate_inside_select_case => {
        "program t\ninteger :: code = 2\nselect case (code)\ncase (1)\nprint *, 'one'\ncase (2)\nassociate (label => 'two')\nprint *, label\nend associate\ncase default\nprint *, 'other'\nend select\nend program t\n",
        ["two"]
    };

    associate_block_then_outer_print => {
        "program t\ninteger :: val = 9\nassociate (v => val)\nprint *, v\nend associate\nprint *, val\nend program t\n",
        ["9", "9"]
    };

    // ── Multiple names in one associate ──────────────────────────────

    associate_multi_three_scalars => {
        "program t\ninteger :: a = 1, b = 2, c = 3\nassociate (x => a, y => b, z => c)\nprint *, x + y + z\nend associate\nend program t\n",
        ["6"]
    };

    associate_multi_expr_and_var => {
        "program t\ninteger :: n = 5\nassociate (double => n * 2, orig => n)\nprint *, double + orig\nend associate\nend program t\n",
        ["15"]
    };

    associate_multi_dtype_fields => {
        "program t\ntype :: Rect\ninteger :: w, h\nend type Rect\ntype(Rect) :: r\nr%w = 4\nr%h = 5\nassociate (width => r%w, height => r%h)\nprint *, width * height\nend associate\nend program t\n",
        ["20"]
    };

    // ── Module and procedure scope ───────────────────────────────────

    associate_in_internal_subroutine => {
        "program t\ncall show()\ncontains\nsubroutine show()\ninteger :: k = 11\nassociate (alias => k)\nprint *, alias\nend associate\nend subroutine show\nend program t\n",
        ["11"]
    };

    associate_in_function_result => {
        "program t\nprint *, double_it(6)\ncontains\ninteger function double_it(n) result(r)\ninteger, intent(in) :: n\nassociate (twice => n * 2)\nr = twice\nend associate\nend function double_it\nend program t\n",
        ["12"]
    };

    associate_module_variable => {
        "module amod\ninteger :: shared = 25\ncontains\nsubroutine peek()\nassociate (s => shared)\nprint *, s\nend associate\nend subroutine peek\nend module amod\nprogram t\nuse amod\ncall peek()\nend program t\n",
        ["25"]
    };

    // ── Complex and mixed-type edge cases ────────────────────────────

    associate_complex_part => {
        "program t\ncomplex :: z = (3.0, 4.0)\nassociate (re => real(z), im => aimag(z))\nprint *, int(re + im)\nend associate\nend program t\n",
        ["7"]
    };

    associate_mixed_int_real_expr => {
        "program t\ninteger :: i = 5\nreal :: r = 2.5\nassociate (mix => real(i) + r)\nprint *, int(mix)\nend associate\nend program t\n",
        ["7"]
    };

    associate_index_from_expression => {
        "program t\ninteger :: a(10)\na = [(i, i = 1, 10)]\ninteger :: idx = 4\nassociate (elem => a(idx + 1))\nprint *, elem\nend associate\nend program t\n",
        ["5"]
    };

    associate_string_index_char => {
        "program t\ncharacter(len=5) :: s = 'abcde'\nassociate (ch => s(3:3))\nprint *, ch\nend associate\nend program t\n",
        ["c"]
    };

    associate_logical_not_expr => {
        "program t\nlogical :: p = .false.\nassociate (q => .not. p)\nprint *, q\nend associate\nend program t\n",
        ["true"]
    };

    associate_comparison_expr => {
        "program t\ninteger :: a = 10, b = 20\nassociate (less => a < b)\nprint *, less\nend associate\nend program t\n",
        ["true"]
    };

    associate_ior_bitwise => {
        "program t\ninteger :: a = 5, b = 3\nassociate (bits => ior(a, b))\nprint *, bits\nend associate\nend program t\n",
        ["7"]
    };

    associate_ieor_bitwise => {
        "program t\ninteger :: a = 5, b = 3\nassociate (bits => ieor(a, b))\nprint *, bits\nend associate\nend program t\n",
        ["6"]
    };

    associate_iand_bitwise => {
        "program t\ninteger :: a = 5, b = 3\nassociate (bits => iand(a, b))\nprint *, bits\nend associate\nend program t\n",
        ["1"]
    };

    associate_shift_left => {
        "program t\ninteger :: v = 3\nassociate (shifted => ishft(v, 1))\nprint *, shifted\nend associate\nend program t\n",
        ["6"]
    };

    associate_merge_ternary => {
        "program t\ninteger :: a = 10, b = 20\nlogical :: pick = .true.\nassociate (chosen => merge(a, b, pick))\nprint *, chosen\nend associate\nend program t\n",
        ["10"]
    };
}
